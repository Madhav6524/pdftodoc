import os, io, base64, json, time, traceback
from openai import OpenAI
import openai

AI_MODEL = "gpt-4o"

# ──────────────────────────────────────────────────────────────────────────────
# Internal helpers
# ──────────────────────────────────────────────────────────────────────────────

def _resize_if_needed(image_bytes, media_type):
    """Resize image if >5MB. Returns (bytes, scale_x, scale_y)."""
    scale_x = scale_y = 1.0
    if len(image_bytes) <= 5 * 1024 * 1024:
        return image_bytes, scale_x, scale_y
    try:
        from PIL import Image
        pil = Image.open(io.BytesIO(image_bytes))
        ow, oh = pil.size
        pil = pil.resize((ow // 2, oh // 2), Image.LANCZOS)
        buf = io.BytesIO()
        fmt = "JPEG" if media_type == "image/jpeg" else "PNG"
        pil.save(buf, format=fmt)
        scale_x = ow / (ow // 2)
        scale_y = oh / (oh // 2)
        print(f"[AI] Resized image {ow}x{oh} → {ow//2}x{oh//2}")
        return buf.getvalue(), scale_x, scale_y
    except Exception as e:
        print(f"[AI] Resize failed: {e}")
        return image_bytes, scale_x, scale_y


def _call_gpt4o(client, data_uri, prompt, attempt=0):
    """Send one Vision request and return raw text response."""
    response = client.chat.completions.create(
        model=AI_MODEL,
        messages=[{
            "role": "user",
            "content": [
                {"type": "text", "text": prompt},
                {"type": "image_url", "image_url": {"url": data_uri, "detail": "high"}}
            ]
        }],
        max_tokens=1024,
    )
    return response.choices[0].message.content.strip()


def _parse_json(raw):
    """Strip markdown fences and parse JSON."""
    text = raw.strip()
    # Strip ```json ... ``` or ``` ... ```
    if text.startswith("```"):
        text = text.split("```")[1]
        if text.startswith("json"):
            text = text[4:]
    text = text.strip().strip("`").strip()
    return json.loads(text)


# ──────────────────────────────────────────────────────────────────────────────
# Main public function
# ──────────────────────────────────────────────────────────────────────────────

def find_text_with_ai(image_bytes: bytes, find_word: str, media_type: str = "image/png") -> list:
    """
    Use GPT-4o Vision to locate bboxes of find_word in the image.
    Returns list of {"x","y","width","height"} dicts, or [] on failure.
    """
    api_key = os.environ.get("OPENAI_API_KEY")
    if not api_key:
        return []

    send_bytes, scale_x, scale_y = _resize_if_needed(image_bytes, media_type)
    b64      = base64.standard_b64encode(send_bytes).decode("utf-8")
    data_uri = f"data:{media_type};base64,{b64}"

    prompt = (
        f"Find all occurrences of the text '{find_word}' in this image.\n"
        f"Return ONLY a valid JSON array, no markdown, no explanation.\n"
        f'Format: [{{"text": "found_text", "x": 10, "y": 20, "width": 150, "height": 18}}]\n'
        f"x, y = top-left corner in pixels. width, height = bounding box size.\n"
        f"If not found return empty array: []"
    )

    client = OpenAI(api_key=api_key)

    for attempt in range(2):
        try:
            raw    = _call_gpt4o(client, data_uri, prompt, attempt)
            bboxes = _parse_json(raw)
            if not isinstance(bboxes, list):
                return []
            result = []
            for bb in bboxes:
                result.append({
                    "x":      int(bb.get("x", 0)      * scale_x),
                    "y":      int(bb.get("y", 0)      * scale_y),
                    "width":  int(bb.get("width", 50)  * scale_x),
                    "height": int(bb.get("height", 20) * scale_y),
                })
            print(f"[AI] GPT-4o found {len(result)} bbox(es) for '{find_word}'")
            return result

        except openai.RateLimitError:
            if attempt == 0:
                print("[AI] Rate limited — retrying in 2s…")
                time.sleep(2)
            else:
                return []
        except json.JSONDecodeError as e:
            print(f"[AI] Invalid JSON: {e} | raw={raw[:200]}")
            return []
        except Exception as e:
            print(f"[AI] Error: {e}")
            return []
    return []


def analyse_image_with_ai(image_bytes: bytes, find_word: str,
                           media_type: str = "image/png") -> dict:
    """
    Ask GPT-4o to analyse the image and return:
      - font_style: "mono" | "sans" | "serif"
      - font_weight: "normal" | "bold"
      - text_color: [R, G, B]
      - bg_color:   [R, G, B]
      - is_dark_bg: bool
      - occurrences: [{x, y, width, height}, ...]

    Returns {} on failure (caller falls back to heuristics).
    """
    api_key = os.environ.get("OPENAI_API_KEY")
    if not api_key:
        return {}

    send_bytes, scale_x, scale_y = _resize_if_needed(image_bytes, media_type)
    b64      = base64.standard_b64encode(send_bytes).decode("utf-8")
    data_uri = f"data:{media_type};base64,{b64}"

    prompt = (
        f"Analyze this image carefully. Find all occurrences of the text '{find_word}'.\n\n"
        f"Return ONLY a valid JSON object (no markdown, no explanation):\n"
        f'{{\n'
        f'  "font_style": "mono",\n'
        f'  "font_weight": "normal",\n'
        f'  "text_color": [255, 255, 255],\n'
        f'  "bg_color": [0, 0, 0],\n'
        f'  "is_dark_bg": true,\n'
        f'  "occurrences": [{{"x": 10, "y": 20, "width": 150, "height": 18}}]\n'
        f'}}\n\n'
        f"Rules:\n"
        f"- font_style: 'mono' for terminal/code/fixed-width fonts, 'serif' for Times/Georgia, 'sans' for Arial/Helvetica\n"
        f"- text_color: the RGB color of the '{find_word}' text pixels (sample actual color)\n"
        f"- bg_color: the RGB color of the background behind the text\n"
        f"- is_dark_bg: true if the background luminance is dark (like a terminal)\n"
        f"- occurrences: pixel bounding boxes of each '{find_word}' instance (empty array [] if not found)\n"
        f"- If text not found, set occurrences to []"
    )

    client = OpenAI(api_key=api_key)

    for attempt in range(2):
        try:
            raw  = _call_gpt4o(client, data_uri, prompt, attempt)
            data = _parse_json(raw)
            if not isinstance(data, dict):
                return {}

            # Scale bboxes back if image was resized
            occs = []
            for bb in data.get("occurrences", []):
                occs.append({
                    "x":      int(bb.get("x", 0)      * scale_x),
                    "y":      int(bb.get("y", 0)      * scale_y),
                    "width":  int(bb.get("width", 50)  * scale_x),
                    "height": int(bb.get("height", 20) * scale_y),
                })
            data["occurrences"] = occs

            print(f"[AI] Image analysis: style={data.get('font_style')} "
                  f"dark={data.get('is_dark_bg')} "
                  f"found={len(occs)} bbox(es)")
            return data

        except openai.RateLimitError:
            if attempt == 0:
                print("[AI] Rate limited — retrying in 2s…")
                time.sleep(2)
            else:
                return {}
        except json.JSONDecodeError as e:
            print(f"[AI] Invalid JSON: {e} | raw={raw[:300]}")
            return {}
        except Exception as e:
            print(f"[AI] analyse_image_with_ai error: {e}")
            traceback.print_exc()
            return {}
    return {}
