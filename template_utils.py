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

def get_first_name(full_name: str) -> str:
    """
    Extract the first name from a full name string.
    Properly handles Arabic compound names:
    - Names starting with: عبد, أبو, ابو, أم, ام, ابن, بنت (e.g. عبد الرحمن, أبو بكر)
    - Names ending with الدين / الدين (e.g. سيف الدين, نور الدين, صلاح الدين)
    - Names ending with الله (e.g. منة الله, هبة الله, آية الله)
    - فاطمة الزهراء / فاطمه الزهراء
    - Single names or English names (e.g. Ahmed Ali -> Ahmed).
    """
    if not full_name:
        return ""
    cleaned = str(full_name).replace("\xa0", " ").strip()
    words = cleaned.split()
    if not words:
        return ""

    first = words[0]
    if len(words) > 1:
        # 1. Classical prefixes
        if first in ["عبد", "أبو", "ابو", "أم", "ام", "ابن", "بنت"]:
            return first + " " + words[1]

        # 2. Compound suffixes with "الدين"
        second = words[1].replace("ّ", "")
        if second in ["الدين", "الدّين"] and first in [
            "سيف", "نور", "صلاح", "علاء", "ضياء", "حسام", "بهاء", "جمال",
            "تقي", "عماد", "عز", "بدر", "محي", "محيي", "شمس", "نجم", "سراج"
        ]:
            return first + " " + words[1]

        # 3. Compound with "الله"
        if second == "الله" and first in [
            "منة", "منه", "هبة", "هبه", "آية", "اية", "فضل", "نصر", "سعد", "جار"
        ]:
            return first + " " + words[1]

        # 4. فاطمة الزهراء
        if first in ["فاطمة", "فاطمه"] and second in ["الزهراء", "الزهره"]:
            return first + " " + words[1]

    return first

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

    def _resolve_var(key: str) -> str:
        k = key.strip()
        if not k:
            return ""
        if k in context:
            val = context[k]
            return str(val) if val is not None else ""
        # Case-insensitive direct match
        k_cf = k.casefold()
        for ctx_k, ctx_v in context.items():
            if str(ctx_k).casefold() == k_cf and ctx_v is not None:
                return str(ctx_v)
        # Normalized alias matching (strip spaces, underscores, dashes)
        norm_k = re.sub(r"[\s_\-]+", "", k_cf)
        if norm_k in (
            "name", "arabicname", "firstname", "studentname",
            "first_name", "arabic_name", "full_name", "fullname",
            "الاسم", "الاسمبالعربي", "الاسمالاول", "اسمالطالب", "اسمالمتدرب"
        ):
            for candidate in (
                "name", "Arabic Name", "arabic_name", "arabic name",
                "first_name", "firstname", "الاسم", "الاسم بالعربي"
            ):
                if candidate in context and context[candidate]:
                    return str(context[candidate])
        if norm_k in (
            "phone", "phonenumber", "mobile", "mobilenumber",
            "whatsapp", "whatsappnumber", "رقمالهاتف", "رقمالموبايل",
            "الموبايل", "الهاتف"
        ):
            for candidate in ("phone", "Phone", "mobile", "Mobile", "raw_phone"):
                if candidate in context and context[candidate]:
                    return str(context[candidate])
        return ""

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
            val = _resolve_var(var_name)
            return val if val else fallback.strip()
            
        # Standard variable: {var_name}
        return _resolve_var(inner)

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
