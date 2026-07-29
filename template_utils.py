import re

def normalize_gender(gender_value: str) -> str:
    """Normalize the gender value to 'male' or 'female'."""
    if not gender_value:
        return ""
    text = str(gender_value).strip().casefold()
    if text in ("female", "f", "أنثى", "انثى", "بنت"):
        return "female"
    if text in ("male", "m", "ذكر", "ولد"):
        return "male"
    return ""

def render_template(template_text: str, context: dict, gender: str = "") -> str:
    """
    Render a template with variables and gender conditionals.
    
    Features:
    1. Standard variables: {name} -> context.get("name")
    2. Fallback variables: {name|Coordinator} -> context.get("name") or "Coordinator"
    3. Gender conditionals: {g:male_word:female_word} -> female_word if gender == 'female' else male_word
    """
    if not template_text:
        return ""

    norm_gender = normalize_gender(gender)

    def replacer(match):
        full_match = match.group(0)
        inner = match.group(1)
        
        # Check for gender conditional: {g:male:female}
        if inner.startswith("g:"):
            parts = inner[2:].split(":")
            if len(parts) == 2:
                male_text, female_text = parts
                return female_text if norm_gender == "female" else male_text
            # If invalid format, just return the full match
            return full_match
            
        # Check for fallback: {var_name|fallback_value}
        if "|" in inner:
            var_name, fallback = inner.split("|", 1)
            val = context.get(var_name.strip())
            return str(val) if val else fallback.strip()
            
        # Standard variable: {var_name}
        var_name = inner.strip()
        val = context.get(var_name)
        return str(val) if val else ""

    # Match anything inside {}
    rendered = re.sub(r"\{([^{}]+)\}", replacer, template_text)
    
    # Optional: apply any legacy heuristics if they were used 
    # (e.g. replacing بيك to بيكي for female, if not using the {g::} syntax)
    if norm_gender == "female":
        rendered = re.sub(r"بيك(?!ي)", "بيكي", rendered)
        rendered = re.sub(r"عليك(?![مين])", "عليكي", rendered)
        rendered = re.sub(r"(?<!ع)ليك(?!ي)", "ليكي", rendered)
        rendered = re.sub(r"معاك(?!ي)", "معاكي", rendered)
        rendered = re.sub(r"(?<![اأإآء-ي])بك(?![اأإآء-ي])", "بكي", rendered)
        rendered = re.sub(r"(?<![اأإآء-ي])لك(?![اأإآء-ي])", "لكي", rendered)
        rendered = re.sub(r"متنساش", "متنسيش", rendered)
        rendered = re.sub(r"تفتح ", "تفتحي ", rendered)
        rendered = re.sub(r"بلغني", "بلغيني", rendered)
        rendered = re.sub(r"كلمني", "كلميني", rendered)

    return rendered
