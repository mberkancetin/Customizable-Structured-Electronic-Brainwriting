# SPDX-FileCopyrightText: 2025 Mahmut Berkan Çetin <m.berkancetin@gmail.com> & Selim Gündüz <sgunduz@atu.edu.tr>
#
# SPDX-License-Identifier: MIT


import streamlit as st
import json
import os
import re
from copy import deepcopy

# Import all structures from default_js_config.py
from default_js_config import (
    DEFAULT_TEXT_KEYS,
    DEFAULT_GLOBAL_VARIABLES_STRUCTURE,
    DEFAULT_MODERATOR_VARIABLES_STRUCTURE,
    DEFAULT_LANDINGPAGE_VARIABLES_TEXT_KEYS,
    DEFAULT_ANALYSIS_VARIABLES_TEXT_KEYS,
    DEFAULT_IMAGE_URL,
    DEFAULT_COLAB_GITHUB_URL,
    DEFAULT_NUM_PARTICIPANTS,
    DEFAULT_NUM_ROUNDS,
    DEFAULT_NUM_IDEAS
)

SUPPORTED_SWAP_ALGORITHMS = [
    "CascadingRoundRobin",
    "InterleavedSweepSwap"
]

# --- Configuration for Internationalization (i18n) ---
LOCALES_DIR = "locales"
DEFAULT_LANGUAGE_CODE = "en"
SUPPORTED_LANGUAGES = {
    "en": "English",
    "tr": "Türkçe",
    "de": "Deutsch",
    "es": "Español",
    "fr": "Français"
}

GOOGLE_TRANSLATE_LANGUAGES = {
    'af': 'Afrikaans', 'ak': 'Akan', 'am': 'Amharic', 'ar': 'Arabic', 'as': 'Assamese',
    'ay': 'Aymara', 'az': 'Azerbaijani', 'be': 'Belarusian', 'bg': 'Bulgarian',
    'bho': 'Bhojpuri', 'bm': 'Bambara', 'bn': 'Bengali', 'bs': 'Bosnian', 'ca': 'Catalan',
    'ceb': 'Cebuano', 'ckb': 'Kurdish (Sorani)', 'co': 'Corsican', 'cs': 'Czech',
    'cy': 'Welsh', 'da': 'Danish', 'de': 'German', 'doi': 'Dogri', 'dv': 'Dhivehi',
    'ee': 'Ewe', 'el': 'Greek', 'en': 'English', 'eo': 'Esperanto', 'es': 'Spanish',
    'et': 'Estonian', 'eu': 'Basque', 'fa': 'Persian', 'fi': 'Finnish', 'fr': 'French',
    'fy': 'Frisian', 'ga': 'Irish', 'gd': 'Scots Gaelic', 'gl': 'Galician', 'gn': 'Guarani',
    'gom': 'Goan Konkani', 'gu': 'Gujarati', 'ha': 'Hausa', 'haw': 'Hawaiian', 'he': 'Hebrew',
    'hi': 'Hindi', 'hmn': 'Hmong', 'hr': 'Croatian', 'ht': 'Haitian Creole', 'hu': 'Hungarian',
    'hy': 'Armenian', 'id': 'Indonesian', 'ig': 'Igbo', 'ilo': 'Ilocano', 'is': 'Icelandic',
    'it': 'Italian', 'iw': 'Hebrew (Legacy)', 'ja': 'Japanese', 'jw': 'Javanese', 'ka': 'Georgian',
    'kk': 'Kazakh', 'km': 'Khmer', 'kn': 'Kannada', 'ko': 'Korean', 'kri': 'Krio',
    'ku': 'Kurdish (Kurmanji)', 'ky': 'Kyrgyz', 'la': 'Latin', 'lb': 'Luxembourgish', 'lg': 'Ganda',
    'ln': 'Lingala', 'lo': 'Lao', 'lt': 'Lithuanian', 'lus': 'Mizo (Lushai)', 'lv': 'Latvian',
    'mg': 'Malagasy', 'mi': 'Maori', 'mk': 'Macedonian', 'ml': 'Malayalam', 'mn': 'Mongolian',
    'mni-Mtei': 'Meiteilon (Manipuri)', 'mr': 'Marathi', 'ms': 'Malay', 'mt': 'Maltese',
    'my': 'Myanmar (Burmese)', 'ne': 'Nepali', 'nl': 'Dutch', 'no': 'Norwegian',
    'nso': 'Northern Sotho', 'ny': 'Nyanja (Chichewa)', 'om': 'Oromo', 'or': 'Oriya (Odia)',
    'pa': 'Punjabi', 'pl': 'Polish', 'ps': 'Pashto', 'pt': 'Portuguese', 'qu': 'Quechua',
    'ro': 'Romanian', 'ru': 'Russian', 'rw': 'Kinyarwanda', 'sa': 'Sanskrit', 'sd': 'Sindhi',
    'si': 'Sinhala', 'sk': 'Slovak', 'sl': 'Slovenian', 'sm': 'Samoan', 'sn': 'Shona',
    'so': 'Somali', 'sq': 'Albanian', 'sr': 'Serbian (Cyrillic)', 'sr-Latn': 'Serbian (Latin)',
    'st': 'Southern Sotho', 'su': 'Sundanese', 'sv': 'Swedish', 'sw': 'Swahili', 'ta': 'Tamil',
    'te': 'Telugu', 'tg': 'Tajik', 'th': 'Thai', 'ti': 'Tigrinya', 'tk': 'Turkmen',
    'tl': 'Filipino (Tagalog)', 'tr': 'Turkish', 'ts': 'Tsonga', 'tt': 'Tatar', 'ug': 'Uyghur',
    'uk': 'Ukrainian', 'ur': 'Urdu', 'uz': 'Uzbek', 'vi': 'Vietnamese', 'xh': 'Xhosa',
    'yi': 'Yiddish', 'yo': 'Yoruba', 'zh': 'Chinese', 'zh-CN': 'Chinese (Simplified)',
    'zh-TW': 'Chinese (Traditional)', 'zu': 'Zulu'
}

LANG_NAME_TO_CODE = {v: k for k, v in GOOGLE_TRANSLATE_LANGUAGES.items()}

# --- Translation Loading Function ---
@st.cache_data # Cache the loaded translations
def load_translations(lang_code):
    filepath = os.path.join(LOCALES_DIR, f"{lang_code}.json")
    try:
        with open(filepath, "r", encoding="utf-8") as f:
            translations = json.load(f)
        return translations
    except FileNotFoundError:
        st.warning(f"Translation file for '{lang_code}' not found. Falling back to default ({DEFAULT_LANGUAGE_CODE}).")
        if lang_code != DEFAULT_LANGUAGE_CODE:
            return load_translations(DEFAULT_LANGUAGE_CODE)
        # If default also not found, return empty dict, error will be shown in main
        return {}
    except json.JSONDecodeError:
        st.error(f"Error decoding JSON from translation file for '{lang_code}'. Please check its format.")
        return {}

def _(key, translations_dict, **kwargs):
    """
    Gets the translated string for a key.
    Falls back to a placeholder if not found, for easy identification of missing translations.
    Supports simple string formatting with kwargs.
    """
    raw_string = translations_dict.get(key)
    if raw_string is None:
        # st.warning(_("error_missing_key_format", translations_dict, key=key)) # Avoid recursion if translations_dict is bad
        return f"<{key}_MISSING_IN_LOCALE>"

    try:
        if kwargs: # Only format if arguments are provided
            return raw_string.format(**kwargs)
        return raw_string # Return raw string if no kwargs (e.g. if string contains {} not meant for formatting here)
    except KeyError as e:
        st.warning(_("error_generic_format", translations_dict, key=key, e=e, raw_string=raw_string, args=kwargs))
        return raw_string # Fallback in case of formatting error
    except Exception as e:
        st.warning(f"General formatting error for key '{key}': {e}. Raw: '{raw_string}'")
        return raw_string

# --- Helper Functions ---
def sanitize_sheet_name(name_str):
    if not isinstance(name_str, str): name_str = str(name_str)
    name_str = name_str.replace(" ", "_")
    name_str = re.sub(r'[^\w_]+', '', name_str)
    return name_str if name_str else "Sheet1"

def apply_participant_edits():
    """
    This is the CORRECT callback. It takes the dictionary of edits
    from st.session_state.participant_editor and applies them to the
    master DataFrame in st.session_state.participants_df.
    """
    # Get the dictionary of edits from the session state
    edits = st.session_state.participant_editor

    # Make a copy of the current DataFrame to modify
    df_to_edit = st.session_state.participants_df.copy()

    # Apply the edits
    for row_index, changed_data in edits["edited_rows"].items():
        # The index in the data editor starts from 1, but pandas index starts from 0
        # so we need to be careful if we are using iloc. Let's use loc which is safer.
        # The editor returns the original index, so we can use `loc`.
        for col_name, new_value in changed_data.items():
            df_to_edit.loc[row_index + 1, col_name] = new_value

    # Update the master DataFrame in session state with the newly edited one
    st.session_state.participants_df = df_to_edit

def get_translated_text_from_key(text_key_from_config, translations_dict, default_val_if_missing=""):
    """Helper to get translated text using a known key from DEFAULT_TEXT_KEYS, then from locale files"""
    actual_translation_key = DEFAULT_TEXT_KEYS.get(text_key_from_config, text_key_from_config)
    return translations_dict.get(actual_translation_key, default_val_if_missing)

def is_valid_image_url(url):
    """Basic validation for image URLs."""
    if not url:
        return True  # Empty URL is valid (no image)
    return url.lower().endswith(('.png', '.jpg', '.jpeg', '.gif'))

def get_image_preview_html(image_url):
    """Generate HTML for image preview."""
    if not image_url:
        return ""
    return f"""
    <div style="margin: 10px 0; padding: 10px; border: 1px solid #ddd; border-radius: 5px; text-align: center;">
        <img src="{image_url}" style="max-width: 200px; max-height: 150px;">
    </div>
    """

def initialize_session_state(translations):
    """Initializes or re-initializes session state with translated default values."""

    # Flag to ensure this logic runs only once per language load or on first app load
    # It compares current app language with the language for which defaults were last loaded
    last_loaded_lang = st.session_state.get("last_loaded_lang_for_defaults", None)
    is_first_app_load = "initial_config_set" not in st.session_state

    if is_first_app_load or last_loaded_lang != st.session_state.app_lang:
        # --- Set numeric/structural defaults first if it's the very first load ---
        if is_first_app_load:
            st.session_state.num_participants = DEFAULT_NUM_PARTICIPANTS
            st.session_state.num_rounds = DEFAULT_NUM_ROUNDS
            st.session_state.num_ideas = DEFAULT_NUM_IDEAS
            st.session_state.customize_participant_names = False
            st.session_state.generate_landing_page = True
            st.session_state.imageUrl = DEFAULT_IMAGE_URL # Non-translatable
            st.session_state.use_logo = True if st.session_state.imageUrl else False
            st.session_state.colabGitHubUrl = DEFAULT_COLAB_GITHUB_URL # Non-translatable
            st.session_state.advanced_settings = False

        # --- GLOBAL_VARIABLES ---
        gv = deepcopy(DEFAULT_GLOBAL_VARIABLES_STRUCTURE) # Start with structure
        gv["SESSION_FOCUS"] = get_translated_text_from_key("SESSION_FOCUS", translations)
        gv["LANDING_SHEET"] = get_translated_text_from_key("LANDING_SHEET", translations) # Assumes sheet names are translated
        gv["MODERATOR_SHEET"] = get_translated_text_from_key("MODERATOR_SHEET", translations) # Assumes sheet names are translated
        gv["SESSION_LANGUAGE"] = get_translated_text_from_key("SESSION_LANGUAGE", translations) # Assumes sheet names are translated
        gv["IDEA_SWAP_ALGORITHM"] = get_translated_text_from_key("IDEA_SWAP_ALGORITHM", translations)
        gv["PARTICIPANT_SHEET_LANGUAGE"] = [get_translated_text_from_key("SESSION_LANGUAGE", translations) for _ in range(st.session_state.num_participants)]
        gv["TIME_LEFT"] = get_translated_text_from_key("TIME_LEFT", translations)
        gv["MINS_LEFT"] = get_translated_text_from_key("MINS_LEFT", translations)
        gv["ONE_MIN_LEFT"] = get_translated_text_from_key("ONE_MIN_LEFT", translations)
        gv["TIME_IS_UP"] = get_translated_text_from_key("TIME_IS_UP", translations)
        gv["FOCUS"] = [
            get_translated_text_from_key("FOCUS_ITEM_1", translations),
            get_translated_text_from_key("FOCUS_ITEM_2", translations),
            get_translated_text_from_key("FOCUS_ITEM_3", translations)
        ]
        gv["CHECK_IDEAS"] = get_translated_text_from_key("CHECK_IDEAS", translations)
        gv["ROUND_CHANGE"] = [
            get_translated_text_from_key("ROUND_CHANGE_MSG1", translations),
            get_translated_text_from_key("ROUND_CHANGE_MSG2", translations)
        ]
        gv["SESSION_COMPLETE"] = get_translated_text_from_key("SESSION_COMPLETE", translations)
        gv["STARTING"] = get_translated_text_from_key("STARTING", translations)
        gv["STOPPED"] = get_translated_text_from_key("STOPPED", translations)
        gv["MINUTES"] = DEFAULT_GLOBAL_VARIABLES_STRUCTURE["MINUTES"] # Non-translated from structure

        participant_prefix = get_translated_text_from_key("PARTICIPANT_PREFIX", translations)
        gv["PARTICIPANT"] = [participant_prefix.format(no=i+1) for i in range(st.session_state.num_participants)]
        round_prefix = get_translated_text_from_key("ROUND_PREFIX", translations)
        gv["ROUNDS"] = [round_prefix.format(no=i+1) for i in range(st.session_state.num_rounds)]
        idea_prefix = get_translated_text_from_key("IDEA_PREFIX", translations)
        gv["IDEAS"] = [idea_prefix.format(no=i+1) for i in range(st.session_state.num_ideas)]
        st.session_state.GLOBAL_VARIABLES = gv

        # --- MODERATOR_VARIABLES ---
        mv = deepcopy(DEFAULT_MODERATOR_VARIABLES_STRUCTURE)
        mv["MENU"]["Tools"] = get_translated_text_from_key("MOD_TOOLS", translations)
        mv["MENU"]["LandingPage"] = get_translated_text_from_key("MOD_LANDING_PAGE", translations)
        mv["MENU"]["SessionElements"] = get_translated_text_from_key("MOD_SESSION_ELEMENTS", translations)
        mv["MENU"]["Start"] = get_translated_text_from_key("MOD_START", translations)
        mv["MENU"]["SubmitNext"] = get_translated_text_from_key("MOD_SUBMIT_NEXT", translations)
        mv["MENU"]["PrepareData"] = get_translated_text_from_key("MOD_PREPARE_DATA", translations)
        mv["MENU"]["ColabEnvironment"] = get_translated_text_from_key("MOD_COLAB_ENV", translations)
        mv["SESSION_START_TEMPLATE"] = get_translated_text_from_key("MOD_SESSION_START_TEMPLATE", translations)
        mv["ROUND_END_PHRASE"] = get_translated_text_from_key("MOD_ROUND_END_PHRASE", translations)
        mv["CURRENT_ROUND"] = get_translated_text_from_key("MOD_CURRENT_ROUND", translations)
        mv["DATA_PREP"]["SheetName"] = get_translated_text_from_key("MOD_DATA_PREP_SHEET_NAME", translations) # Assumes translated
        mv["DATA_PREP"]["IdeaRawColumn"] = get_translated_text_from_key("MOD_DATA_PREP_IDEA_RAW_COL", translations) # Assumes translated
        mv["DATA_PREP"]["TranslateColumn"] = get_translated_text_from_key("MOD_DATA_PREP_TRANSLATE_COL", translations) # Assumes translated
        mv["DATA_PREP"]["IdeaID"] = get_translated_text_from_key("MOD_DATA_PREP_IDEAID_COL", translations)
        mv["DATA_PREP"]["IdeaTimestamp"] = get_translated_text_from_key("MOD_DATA_PREP_IDEATIMESTAMP_COL", translations)
        mv["DATA_PREP"]["IdeaRoundStartTimestamp"] = get_translated_text_from_key("MOD_DATA_PREP_ROUNDTIMESTAMP_COL", translations)
        mv["DATA_PREP"]["ManualCategorization"] = get_translated_text_from_key("MOD_DATA_PREP_MANUAL_CAT", translations) # Assumes translated
        # Non-translated parts from structure
        mv["DATA_PREP"]["SessionLanguage"] = DEFAULT_MODERATOR_VARIABLES_STRUCTURE["DATA_PREP"]["SessionLanguage"]
        mv["DATA_PREP"]["TranslatedLanguage"] = DEFAULT_MODERATOR_VARIABLES_STRUCTURE["DATA_PREP"]["TranslatedLanguage"]
        mv["COLORS"] = deepcopy(DEFAULT_MODERATOR_VARIABLES_STRUCTURE["COLORS"])
        mv["START_TIMER"] = get_translated_text_from_key("MOD_START_TIMER", translations)
        st.session_state.MODERATOR_VARIABLES = mv

        # --- LANDINGPAGE_VARIABLES_TEXTS ---
        lp_texts = {}
        for key_in_code, text_key_for_translation_lookup in DEFAULT_LANDINGPAGE_VARIABLES_TEXT_KEYS.items():
            # text_key_for_translation_lookup is like "LP_GREETING_MESSAGE_TEMPLATE"
            # This key is then used in get_translated_text_from_key to find "default_lp_greeting_message_template" in DEFAULT_TEXT_KEYS
            # and then "default_lp_greeting_message_template" is used to get the actual text from `translations`
            lp_texts[key_in_code] = get_translated_text_from_key(text_key_for_translation_lookup, translations)
        st.session_state.LANDINGPAGE_VARIABLES_TEXTS = lp_texts

        # --- ANALYSIS_VARIABLES_TEXTS ---
        an_texts = {}
        for key_in_code, text_key_for_translation_lookup in DEFAULT_ANALYSIS_VARIABLES_TEXT_KEYS.items():
            an_texts[key_in_code] = get_translated_text_from_key(text_key_for_translation_lookup, translations)
        st.session_state.ANALYSIS_VARIABLES_TEXTS = an_texts

        st.session_state.initial_config_set = True
        st.session_state.last_loaded_lang_for_defaults = st.session_state.app_lang
        if not is_first_app_load: # If it's a language change, not first load
             st.toast(_("default_config_loaded_msg", translations))


def generate_js_from_state(translations_for_error_msg):
    try:
        if not st.session_state.advanced_settings:
            with open("template.js", "r", encoding="utf-8") as f:
                template_str = f.read()
        else:
            with open("template_multilingual.js", "r", encoding="utf-8") as f:
                template_str = f.read()
    except FileNotFoundError:
        st.error(_("error_template_file", translations_for_error_msg))
        return None

    gv = st.session_state.GLOBAL_VARIABLES
    mv = st.session_state.MODERATOR_VARIABLES
    lp_texts = st.session_state.LANDINGPAGE_VARIABLES_TEXTS
    an_texts = st.session_state.ANALYSIS_VARIABLES_TEXTS

    session_start_full = mv["SESSION_START_TEMPLATE"]
    parts = session_start_full.split("{ideasCount}")
    mod_session_start_part1 = parts[0]
    mod_session_start_part2 = parts[1] if len(parts) > 1 else ""

    # Prepare greeting message with session focus interpolated
    greeting_template = lp_texts.get("GREETING_MESSAGE_TEMPLATE", "")
    populated_greeting = greeting_template.replace("{SESSION_FOCUS}", gv.get("SESSION_FOCUS", ""))

    # Sanitize participant names before dumping to JS array
    raw_participant_list = gv.get("PARTICIPANT", [])
    participant_sheet_language_list = gv.get("PARTICIPANT_SHEET_LANGUAGE", [])
    sanitized_participant_list = [sanitize_sheet_name(name) for name in raw_participant_list]


    replacements = {
        "{{SESSION_FOCUS_JS_STRING}}": json.dumps(gv.get("SESSION_FOCUS", ""), ensure_ascii=False),
        "{{LANDING_SHEET_JS_STRING}}": json.dumps(sanitize_sheet_name(gv.get("LANDING_SHEET", "Welcome")), ensure_ascii=False),
        "{{PARTICIPANT_LIST_AS_JS_ARRAY}}": json.dumps(sanitized_participant_list, ensure_ascii=False),
        "{{MODERATOR_SHEET_JS_STRING}}": json.dumps(sanitize_sheet_name(gv.get("MODERATOR_SHEET", "Moderator")), ensure_ascii=False),
        "{{SESSION_LANGUAGE_JS_STRING}}": json.dumps(gv.get("SESSION_LANGUAGE", ""), ensure_ascii=False),
        "{{IDEA_SWAP_ALGORITHM_JS_STRING}}": json.dumps(gv.get("IDEA_SWAP_ALGORITHM", ""), ensure_ascii=False),
        "{{PARTICIPANT_SHEET_LANGUAGE_LIST_AS_JS_ARRAY}}": json.dumps(participant_sheet_language_list, ensure_ascii=False),
        "{{TIME_LEFT_JS_STRING}}": json.dumps(gv.get("TIME_LEFT", ""), ensure_ascii=False),
        "{{MINS_LEFT_JS_STRING}}": json.dumps(gv.get("MINS_LEFT", ""), ensure_ascii=False),
        "{{ONE_MIN_LEFT_JS_STRING}}": json.dumps(gv.get("ONE_MIN_LEFT", ""), ensure_ascii=False),
        "{{TIME_IS_UP_JS_STRING}}": json.dumps(gv.get("TIME_IS_UP", ""), ensure_ascii=False),
        "{{MINUTES_INT}}": str(gv.get("MINUTES", 5)),
        "{{FOCUS_LIST_AS_JS_ARRAY}}": json.dumps(gv.get("FOCUS", []), ensure_ascii=False),
        "{{ROUNDS_LIST_AS_JS_ARRAY}}": json.dumps(gv.get("ROUNDS", []), ensure_ascii=False),
        "{{IDEAS_LIST_AS_JS_ARRAY}}": json.dumps(gv.get("IDEAS", []), ensure_ascii=False),
        "{{CHECK_IDEAS_JS_STRING}}": json.dumps(gv.get("CHECK_IDEAS", ""), ensure_ascii=False),
        "{{ROUND_CHANGE_LIST_AS_JS_ARRAY}}": json.dumps(gv.get("ROUND_CHANGE", []), ensure_ascii=False),
        "{{SESSION_COMPLETE_JS_STRING}}": json.dumps(gv.get("SESSION_COMPLETE", ""), ensure_ascii=False),
        "{{STARTING_JS_STRING}}": json.dumps(gv.get("STARTING", ""), ensure_ascii=False),
        "{{STOPPED_JS_STRING}}": json.dumps(gv.get("STOPPED", ""), ensure_ascii=False),

        "{{MENU_OBJECT_AS_JS_OBJECT}}": json.dumps(mv.get("MENU", {}), ensure_ascii=False),
        "{{MOD_SESSION_START_TEMPLATE_STRING_PART1}}": json.dumps(mod_session_start_part1, ensure_ascii=False),
        "{{MOD_SESSION_START_TEMPLATE_STRING_PART2}}": json.dumps(mod_session_start_part2, ensure_ascii=False),
        "{{ROUND_END_PHRASE_JS_STRING}}": json.dumps(mv.get("ROUND_END_PHRASE", ""), ensure_ascii=False),
        "{{CURRENT_ROUND_JS_STRING}}": json.dumps(mv.get("CURRENT_ROUND", ""), ensure_ascii=False),
        "{{DATA_PREP_SheetName_JS_STRING}}": json.dumps(sanitize_sheet_name(mv.get("DATA_PREP", {}).get("SheetName", "PrepData")), ensure_ascii=False),
        "{{DATA_PREP_SessionLanguage_JS_STRING}}": json.dumps(mv.get("DATA_PREP", {}).get("SessionLanguage", "auto"), ensure_ascii=False),
        "{{DATA_PREP_IdeaRawColumn_JS_STRING}}": json.dumps(mv.get("DATA_PREP", {}).get("IdeaRawColumn", "RawIdea"), ensure_ascii=False),
        "{{DATA_PREP_TranslatedLanguage_JS_STRING}}": json.dumps(mv.get("DATA_PREP", {}).get("TranslatedLanguage", "en"), ensure_ascii=False),
        "{{DATA_PREP_IdeaID_JS_STRING}}": json.dumps(mv.get("DATA_PREP", {}).get("IdeaID", "IdeaID"), ensure_ascii=False),
        "{{DATA_PREP_IdeaTimestamp_JS_STRING}}": json.dumps(mv.get("DATA_PREP", {}).get("IdeaTimestamp", "IdeaTimestamp"), ensure_ascii=False),
        "{{DATA_PREP_IdeaRoundStartTimestamp_JS_STRING}}": json.dumps(mv.get("DATA_PREP", {}).get("IdeaRoundStartTimestamp", "RoundTimestamp"), ensure_ascii=False),
        "{{DATA_PREP_TranslateColumn_JS_STRING}}": json.dumps(mv.get("DATA_PREP", {}).get("TranslateColumn", "Translation"), ensure_ascii=False),
        "{{DATA_PREP_ManualCategorization_JS_STRING}}": json.dumps(mv.get("DATA_PREP", {}).get("ManualCategorization", "ManualCategories"), ensure_ascii=False),
        "{{COLORS_OBJECT_AS_JS_OBJECT}}": json.dumps(mv.get("COLORS", {}), ensure_ascii=False),
        "{{START_TIMER_OBJECT_AS_JS_OBJECT}}": json.dumps(mv.get("START_TIMER", {}), ensure_ascii=False),

        "{{IMAGE_URL_JS_STRING}}": json.dumps(st.session_state.imageUrl if st.session_state.get("generate_landing_page", False) and st.session_state.get("use_logo", False) else ""),
        "{{COLAB_GITHUB_URL_JS_STRING}}": json.dumps(st.session_state.get("colabGitHubUrl", ""), ensure_ascii=False),

        "{{LP_GREETING_MESSAGE_JS_STRING}}": json.dumps(populated_greeting, ensure_ascii=False),
        "{{LP_SESSION_TITLE_SUFFIX_JS_STRING}}": json.dumps(lp_texts.get("SESSION_TITLE_SUFFIX", ""), ensure_ascii=False),
        "{{LP_PARTICIPANTS_PREFIX_JS_STRING}}": json.dumps(lp_texts.get("PARTICIPANTS_PREFIX", ""), ensure_ascii=False),
        "{{LP_ROUNDS_PREFIX_JS_STRING}}": json.dumps(lp_texts.get("ROUNDS_PREFIX", ""), ensure_ascii=False),
        "{{LP_IDEAS_PREFIX_JS_STRING}}": json.dumps(lp_texts.get("IDEAS_PREFIX", ""), ensure_ascii=False),
        "{{LP_TIME_PREFIX_JS_STRING}}": json.dumps(lp_texts.get("TIME_PREFIX", ""), ensure_ascii=False),
        "{{LP_TOTALS_TEXT_INFIX_JS_STRING}}": json.dumps(lp_texts.get("TOTALS_TEXT_INFIX", ""), ensure_ascii=False),

        "{{ANALYSIS_POPUP_TITLE_JS_STRING}}": json.dumps(an_texts.get("POPUP_TITLE", ""), ensure_ascii=False),
        "{{ANALYSIS_COLAB_POPUP_NEXT_STEPS_JS_STRING}}": json.dumps(an_texts.get("COLAB_POPUP_NEXT_STEPS", ""), ensure_ascii=False),
        "{{ANALYSIS_COLAB_POPUP_STEP1_PART1_JS_STRING}}": json.dumps(an_texts.get("COLAB_POPUP_STEP1_PART1", ""), ensure_ascii=False),
        "{{ANALYSIS_COLAB_POPUP_STEP1_LINK_TEXT_JS_STRING}}": json.dumps(an_texts.get("COLAB_POPUP_STEP1_LINK_TEXT", ""), ensure_ascii=False),
        "{{ANALYSIS_COLAB_POPUP_STEP1_SMALL_TEXT_JS_STRING}}": json.dumps(an_texts.get("COLAB_POPUP_STEP1_SMALL_TEXT", ""), ensure_ascii=False),
        "{{ANALYSIS_COLAB_POPUP_STEP2_JS_STRING}}": json.dumps(an_texts.get("COLAB_POPUP_STEP2", ""), ensure_ascii=False),
        "{{ANALYSIS_COLAB_POPUP_STEP3_JS_STRING}}": json.dumps(an_texts.get("COLAB_POPUP_STEP3", ""), ensure_ascii=False),
    }

    for placeholder, value in replacements.items():
        template_str = template_str.replace(placeholder, value)
    return template_str

# --- Main Application ---
def main():
    # Initialize app_lang first if not set, done before loading any translations
    if "app_lang" not in st.session_state:
        st.session_state.app_lang = DEFAULT_LANGUAGE_CODE

    st.markdown(
        """
        <style>
            [data-testid="stSidebar"] {
                width: 450px !important;
            }
        </style>
        """,
        unsafe_allow_html=True
    )
    # --- Language Selection ---
    # Load translations just for the sidebar elements initially
    sidebar_translations = load_translations(st.session_state.app_lang)
    if not sidebar_translations: # Critical check
        st.error(f"Failed to load translations for '{st.session_state.app_lang}'. App cannot start.")
        return

    st.sidebar.header(_("sidebar_settings_header", sidebar_translations))
    current_lang_index = 0
    try:
        current_lang_index = list(SUPPORTED_LANGUAGES.keys()).index(st.session_state.app_lang)
    except ValueError: pass # Default to 0 if somehow current lang not in keys

    selected_lang_display_name = st.sidebar.selectbox(
        label=_("app_language_label", sidebar_translations),
        options=list(SUPPORTED_LANGUAGES.values()),
        index=current_lang_index,
        key="lang_selector_widget"
    )
    new_lang_code = next(code for code, name in SUPPORTED_LANGUAGES.items() if name == selected_lang_display_name)

    language_has_changed = st.session_state.app_lang != new_lang_code
    if language_has_changed:
        st.session_state.app_lang = new_lang_code

    # Load full T_UI for the selected (or new) language
    T_UI = load_translations(st.session_state.app_lang)
    if not T_UI:
        error_message_key = "critical_error_no_translations"
        error_message = sidebar_translations.get(error_message_key, f"<{error_message_key}_MISSING_IN_LOCALE> Critical error: Main UI translations could not be loaded.")
        st.error(error_message)
        return

    st.session_state.advanced_settings = st.sidebar.toggle(
                _("advanced_settings", T_UI),
                value=False,
                key="fc_toggle_advanced_settings",
                help=_("help_advanced_settings", T_UI)
            )

    # Instructions and tips
    st.sidebar.markdown(_("instructions_and_tips", T_UI))


    # Initialize or Re-initialize session state with translated defaults
    # This is where default values for input fields get set based on T_UI
    initialize_session_state(T_UI)

    if language_has_changed:
        st.rerun() # Rerun NOW to apply new language defaults to all widgets and redraw UI

    # --- Start of the main page UI, using T_UI for all text ---
    st.title(_("app_title", T_UI))
    # st.info(_("author_info", T_UI))

    # --- Configuration Tabs ---
    tab_keys = ["tab_quick_setup", "tab_js_localization", "tab_full_customization"]
    tab_labels = [_(key, T_UI) for key in tab_keys]
    tab_quick, tab_localize, tab_full_customize = st.tabs(tab_labels)

    # --- Quick Setup Tab ---
    with tab_quick:
        st.header(_("tab_quick_setup", T_UI))
        st.session_state.GLOBAL_VARIABLES["SESSION_FOCUS"] = st.text_area(
            _("session_focus_label", T_UI),
            st.session_state.GLOBAL_VARIABLES["SESSION_FOCUS"], key="qs_session_focus",
            help=_("help_session_focus", T_UI)
        )

        c1, c2 = st.columns(2)
        with c1:
            if not st.session_state.advanced_settings:
                # Numeric inputs for counts; these will trigger updates in initialize_session_state if changed
                # Or, we handle list regeneration directly if these numbers change
                num_participants_val = st.number_input(
                    _("num_participants_label", T_UI), min_value=1,
                    value=st.session_state.num_participants, step=1, key="qs_num_participants_input",
                    help=_("help_number_participants", T_UI)
                )
                if num_participants_val != st.session_state.num_participants:
                    st.session_state.num_participants = num_participants_val
                    # Regenerate participant list if not customized
                    if not st.session_state.get("customize_participant_names", False):
                        participant_prefix = get_translated_text_from_key("PARTICIPANT_PREFIX", T_UI)
                        st.session_state.GLOBAL_VARIABLES["PARTICIPANT"] = [participant_prefix.format(no=i+1) for i in range(st.session_state.num_participants)]
                    st.rerun() # Rerun to reflect changes, especially if names are generated

                st.session_state.customize_participant_names = st.toggle(
                    _("customize_participant_names_label", T_UI),
                    value=st.session_state.customize_participant_names,
                    key="fc_toggle_customize_names1",
                    help=_("help_list_participants", T_UI)
                )
                st.caption(_("data_editor_info_dict", T_UI))
                if st.session_state.customize_participant_names:
                    current_names = list(st.session_state.GLOBAL_VARIABLES["PARTICIPANT"])
                    num_p_display = st.session_state.num_participants

                    # Adjust list size if num_participants changed from quick setup
                    if len(current_names) > num_p_display: current_names = current_names[:num_p_display]
                    elif len(current_names) < num_p_display:
                        participant_prefix = get_translated_text_from_key("PARTICIPANT_PREFIX", T_UI)
                        current_names.extend([participant_prefix.format(no=i+1) for i in range(len(current_names), num_p_display)])

                    new_participants = []
                    for i in range(num_p_display):
                        new_participants.append(st.text_input(
                            _("participant_name_label", T_UI, i=i+1),
                            value=current_names[i] if i < len(current_names) else f"{get_translated_text_from_key('PARTICIPANT_PREFIX', T_UI)}{i+1}",
                            key=f"fc_p_name_custom_{i}"
                        ))
                    sanitized_participant_list = [sanitize_sheet_name(name) for name in new_participants]
                    st.session_state.GLOBAL_VARIABLES["PARTICIPANT"] = sanitized_participant_list

                else: # Show generated, non-editable list
                    st.write(pd.DataFrame({_("participant_sheet_names", T_UI): st.session_state.GLOBAL_VARIABLES["PARTICIPANT"]}, index=range(1, st.session_state.num_participants+1)))
            else:
                # Numeric input for participant count
                num_participants_val = st.number_input(
                    _("num_participants_label", T_UI), min_value=1,
                    value=st.session_state.num_participants, step=1, key="qs_num_participants_input",
                    help=_("help_number_participants", T_UI)
                )

                # --- Data Structure Management ---
                df_exists = "participants_df" in st.session_state
                if not df_exists or num_participants_val != st.session_state.num_participants:
                    st.session_state.num_participants = num_participants_val
                    participant_prefix = get_translated_text_from_key("PARTICIPANT_PREFIX", T_UI)

                    # Create or update the DataFrame
                    new_names = [participant_prefix.format(no=i+1) for i in range(st.session_state.num_participants)]
                    new_langs = [GOOGLE_TRANSLATE_LANGUAGES[get_translated_text_from_key("SESSION_LANGUAGE", T_UI)]] * st.session_state.num_participants

                    new_df = pd.DataFrame({
                        "Name": new_names,
                        "Language": new_langs
                    })

                    if df_exists:
                        old_df = st.session_state.participants_df
                        common_rows = min(len(old_df), len(new_df))
                        new_df.iloc[:common_rows] = old_df.iloc[:common_rows]

                    new_df.index = range(1, len(new_df) + 1)

                    st.session_state.participants_df = new_df
                    st.rerun()

                # Toggle for customizing
                st.session_state.customize_participant_names = st.toggle(
                    _("customize_participant_names_label", T_UI),
                    value=st.session_state.customize_participant_names,
                    key="fc_toggle_customize_names1",
                    help=_("help_list_participants", T_UI)
                )

                if st.session_state.customize_participant_names:
                    st.caption(_("data_editor_info_dict", T_UI))

                    st.data_editor(
                        st.session_state.participants_df,
                        num_rows="fixed",
                        use_container_width=True,
                        on_change=apply_participant_edits,
                        column_config={
                            "Name": st.column_config.TextColumn(
                                required=True,
                                label=_("participant_sheet_names", T_UI)
                            ),
                            "Language": st.column_config.SelectboxColumn(
                                required=True,
                                label=_("multilingual_sheet_language", T_UI),
                                options=GOOGLE_TRANSLATE_LANGUAGES.values()
                            )
                        },
                        key="participant_editor"
                    )

                    participant_names = st.session_state.participants_df["Name"].tolist()
                    participant_lang_names = st.session_state.participants_df["Language"].tolist()

                    participant_lang_codes = [LANG_NAME_TO_CODE[name] for name in participant_lang_names]

                    st.session_state.GLOBAL_VARIABLES["PARTICIPANT"] = [sanitize_sheet_name(name) for name in participant_names]
                    st.session_state.GLOBAL_VARIABLES["PARTICIPANT_SHEET_LANGUAGE"] = participant_lang_codes

                else:  # Show the non-editable list
                    display_df = st.session_state.participants_df.copy()
                    display_df.index = range(1, len(display_df) + 1)
                    display_df = display_df.rename(columns={'Name': _("participant_sheet_names", T_UI), 'Language': _("multilingual_sheet_language", T_UI)})
                    st.dataframe(display_df, use_container_width=True)

                    participant_lang_names = st.session_state.participants_df["Language"].tolist()
                    participant_lang_codes = [LANG_NAME_TO_CODE[name] for name in participant_lang_names]

                    st.session_state.GLOBAL_VARIABLES["PARTICIPANT"] = [sanitize_sheet_name(name) for name in st.session_state.participants_df["Name"]]
                    st.session_state.GLOBAL_VARIABLES["PARTICIPANT_SHEET_LANGUAGE"] = participant_lang_codes
        with c2:
            num_rounds_val = st.number_input(
                _("num_rounds_label", T_UI), min_value=1,
                value=st.session_state.num_rounds, step=1, key="qs_num_rounds_input",
                help=_("help_number_rounds", T_UI)
            )
            if num_rounds_val != st.session_state.num_rounds:
                st.session_state.num_rounds = num_rounds_val
                round_prefix = get_translated_text_from_key("ROUND_PREFIX", T_UI)
                st.session_state.GLOBAL_VARIABLES["ROUNDS"] = [round_prefix.format(no=i+1) for i in range(st.session_state.num_rounds)]
                st.rerun()

            num_ideas_val = st.number_input(
                _("num_ideas_label", T_UI), min_value=1,
                value=st.session_state.num_ideas, step=1, key="qs_num_ideas_input",
                help=_("help_number_ideas", T_UI)
            )
            if num_ideas_val != st.session_state.num_ideas:
                st.session_state.num_ideas = num_ideas_val
                idea_prefix = get_translated_text_from_key("IDEA_PREFIX", T_UI)
                st.session_state.GLOBAL_VARIABLES["IDEAS"] = [idea_prefix.format(no=i+1) for i in range(st.session_state.num_ideas)]
                st.rerun()

            st.session_state.GLOBAL_VARIABLES["MINUTES"] = st.number_input(
                _("minutes_per_round_label", T_UI), min_value=1,
                value=st.session_state.GLOBAL_VARIABLES["MINUTES"], step=1, key="qs_minutes",
                help=_("help_minutes_round", T_UI)
            )

            st.session_state.generate_landing_page = st.checkbox(
                _("generate_landing_page_label", T_UI), value=st.session_state.generate_landing_page, key="qs_gen_lp"
            )
            if st.session_state.generate_landing_page:
                st.session_state.use_logo = st.checkbox(
                    _("use_logo_label", T_UI), value=False, key="qs_use_logo"
                )
                if st.session_state.use_logo:
                    new_image_url = st.text_input(
                        _("logo_url_label", T_UI), st.session_state.imageUrl, key="qs_logo_url"
                    )

                    if new_image_url != st.session_state.imageUrl: # Only update if changed
                        st.session_state.imageUrl = new_image_url
                        # No need for st.rerun() just for text input change unless preview depends on it immediately

                    if st.session_state.imageUrl: # Only try to validate/preview if URL is not empty
                        if not is_valid_image_url(st.session_state.imageUrl):
                            st.warning(_("help_logo_landing_page", T_UI))
                        else:
                            # Display image preview
                            preview_html = get_image_preview_html(st.session_state.imageUrl)
                            if preview_html:
                                st.markdown(preview_html, unsafe_allow_html=True)
                    elif new_image_url and not st.session_state.imageUrl: # if user just cleared the URL
                        pass # No warning if URL is empty

    # --- JS Variable Localization Tab ---
    with tab_localize:
        st.header(_("tab_js_localization", T_UI))
        with st.expander(_("rounds_list_header", T_UI)):
            st.caption(_("data_editor_info_list", T_UI))
            st.session_state.GLOBAL_VARIABLES["ROUNDS"] = st.data_editor(
                pd.DataFrame({"Round Names": st.session_state.GLOBAL_VARIABLES["ROUNDS"]}),
                num_rows="dynamic", hide_index=True, key="fc_rounds_editor"
            )["Round Names"].tolist()


        with st.expander(_("ideas_list_header", T_UI)):
            st.caption(_("data_editor_info_list", T_UI))
            st.session_state.GLOBAL_VARIABLES["IDEAS"] = st.data_editor(
                 pd.DataFrame({"Idea Names": st.session_state.GLOBAL_VARIABLES["IDEAS"]}),
                num_rows="dynamic", hide_index=True, key="fc_ideas_editor"
            )["Idea Names"].tolist()

        with st.expander(_("focus_list_label", T_UI)):
             st.caption(_("data_editor_info_list", T_UI))
             st.session_state.GLOBAL_VARIABLES["FOCUS"] = st.data_editor(
                 pd.DataFrame({"Focus Items": st.session_state.GLOBAL_VARIABLES["FOCUS"]}),
                num_rows="dynamic", hide_index=True, key="fc_focus_editor"
            )["Focus Items"].tolist()

        with st.expander(_("round_change_list_label", T_UI)):
             st.caption(_("data_editor_info_list", T_UI))
             st.session_state.GLOBAL_VARIABLES["ROUND_CHANGE"] = st.data_editor(
                 pd.DataFrame({"Messages": st.session_state.GLOBAL_VARIABLES["ROUND_CHANGE"]}),
                num_rows="dynamic", hide_index=True, key="fc_round_change_editor"
            )["Messages"].tolist()

        # GLOBAL_VARIABLES texts
        with st.expander(_("global_vars_texts_header", T_UI)):
            st.session_state.GLOBAL_VARIABLES["TIME_LEFT"] = st.text_input(_("time_left_label", T_UI), st.session_state.GLOBAL_VARIABLES["TIME_LEFT"], key="loc_time_left")
            st.session_state.GLOBAL_VARIABLES["MINS_LEFT"] = st.text_input(_("default_mins_left", T_UI), st.session_state.GLOBAL_VARIABLES["MINS_LEFT"], key="loc_mins_left")
            st.session_state.GLOBAL_VARIABLES["ONE_MIN_LEFT"] = st.text_input(_("default_one_min_left", T_UI), st.session_state.GLOBAL_VARIABLES["ONE_MIN_LEFT"], key="loc_one_min_left")

            st.session_state.GLOBAL_VARIABLES["TIME_IS_UP"] = st.text_input(_("time_is_up_label", T_UI), st.session_state.GLOBAL_VARIABLES["TIME_IS_UP"], key="loc_time_is_up")
            st.session_state.GLOBAL_VARIABLES["CHECK_IDEAS"] = st.text_input(_("check_ideas_label", T_UI), st.session_state.GLOBAL_VARIABLES["CHECK_IDEAS"], key="loc_check_ideas")
            st.session_state.GLOBAL_VARIABLES["SESSION_COMPLETE"] = st.text_input(_("session_complete_label", T_UI), st.session_state.GLOBAL_VARIABLES["SESSION_COMPLETE"], key="loc_session_complete")
            st.session_state.GLOBAL_VARIABLES["STARTING"] = st.text_input(_("starting_label", T_UI), st.session_state.GLOBAL_VARIABLES["STARTING"], key="loc_starting")
            st.session_state.GLOBAL_VARIABLES["STARTING"] = st.text_input(_("stoppped_label", T_UI), st.session_state.GLOBAL_VARIABLES["STARTING"], key="loc_stopped")

            st.session_state.MODERATOR_VARIABLES["SESSION_START_TEMPLATE"] = st.text_area(
                            _("session_start_label", T_UI), st.session_state.MODERATOR_VARIABLES["SESSION_START_TEMPLATE"], height=100, key="loc_session_start"
                        )

        # LANDINGPAGE_VARIABLES_TEXTS
        with st.expander(_("lp_texts_header", T_UI)):
            st.session_state.GLOBAL_VARIABLES["LANDING_SHEET"] = st.text_input(
                _("landing_sheet_label", T_UI), st.session_state.GLOBAL_VARIABLES["LANDING_SHEET"], key="qs_landing_sheet"
            )
            st.session_state.LANDINGPAGE_VARIABLES_TEXTS["GREETING_MESSAGE_TEMPLATE"] = st.text_area(
                _("greeting_msg_label", T_UI), st.session_state.LANDINGPAGE_VARIABLES_TEXTS["GREETING_MESSAGE_TEMPLATE"], height=300, key="loc_lp_greeting"
            )
            st.session_state.LANDINGPAGE_VARIABLES_TEXTS["SESSION_TITLE_SUFFIX"] = st.text_input(_("lp_session_title_suffix_label", T_UI), st.session_state.LANDINGPAGE_VARIABLES_TEXTS["SESSION_TITLE_SUFFIX"], key="loc_lp_title_suffix")
            st.session_state.LANDINGPAGE_VARIABLES_TEXTS["PARTICIPANTS_PREFIX"] = st.text_input(_("lp_participants_prefix_label", T_UI), st.session_state.LANDINGPAGE_VARIABLES_TEXTS["PARTICIPANTS_PREFIX"], key="loc_lp_part_prefix")
            st.session_state.LANDINGPAGE_VARIABLES_TEXTS["ROUNDS_PREFIX"] = st.text_input(_("lp_rounds_prefix_label", T_UI), st.session_state.LANDINGPAGE_VARIABLES_TEXTS["ROUNDS_PREFIX"], key="loc_lp_rounds_prefix")
            st.session_state.LANDINGPAGE_VARIABLES_TEXTS["IDEAS_PREFIX"] = st.text_input(_("lp_ideas_prefix_label", T_UI), st.session_state.LANDINGPAGE_VARIABLES_TEXTS["IDEAS_PREFIX"], key="loc_lp_ideas_prefix")
            st.session_state.LANDINGPAGE_VARIABLES_TEXTS["TIME_PREFIX"] = st.text_input(_("lp_time_prefix_label", T_UI), st.session_state.LANDINGPAGE_VARIABLES_TEXTS["TIME_PREFIX"], key="loc_lp_time_prefix")
            st.session_state.LANDINGPAGE_VARIABLES_TEXTS["TOTALS_TEXT_INFIX"] = st.text_input(_("lp_totals_infix_label", T_UI), st.session_state.LANDINGPAGE_VARIABLES_TEXTS["TOTALS_TEXT_INFIX"], key="loc_lp_totals_infix")

    # --- Advanced Customization Tab ---
    with tab_full_customize:
        st.header(_("tab_full_customization", T_UI))

        with st.expander(_("mod_menu_editor_label", T_UI)):
            st.session_state.GLOBAL_VARIABLES["IDEA_SWAP_ALGORITHM"] = st.selectbox(
                label=_("paper_swap_label", T_UI),
                options=SUPPORTED_SWAP_ALGORITHMS,
                key="swap_selector_dropdown"
            )

            st.session_state.GLOBAL_VARIABLES["MODERATOR_SHEET"] = st.text_input(
                _("moderator_sheet_label", T_UI), st.session_state.GLOBAL_VARIABLES["MODERATOR_SHEET"], key="qs_moderator_sheet"
            )
            st.caption(_("data_editor_info_dict", T_UI))
            if isinstance(st.session_state.MODERATOR_VARIABLES["MENU"], dict):
                st.session_state.MODERATOR_VARIABLES["MENU"] = st.data_editor(
                    st.session_state.MODERATOR_VARIABLES["MENU"], key="fc_menu_editor"
                )
            else: st.warning(_("warning_not_dict", T_UI, variable_name="MODERATOR_VARIABLES.MENU"))
            st.session_state.MODERATOR_VARIABLES["ROUND_END_PHRASE"] = st.text_input(_("round_end_phrase_label", T_UI), st.session_state.MODERATOR_VARIABLES["ROUND_END_PHRASE"], key="loc_round_end_phrase")
            st.session_state.MODERATOR_VARIABLES["CURRENT_ROUND"] = st.text_input(_("current_round_label", T_UI), st.session_state.MODERATOR_VARIABLES["CURRENT_ROUND"], key="loc_current_round")
            st.session_state.ANALYSIS_VARIABLES_TEXTS["POPUP_TITLE"] = st.text_input(_("an_popup_title_label", T_UI), st.session_state.ANALYSIS_VARIABLES_TEXTS["POPUP_TITLE"], key="loc_an_title")
            st.session_state.ANALYSIS_VARIABLES_TEXTS["COLAB_POPUP_NEXT_STEPS"] = st.text_input(_("an_colab_next_steps_label", T_UI), st.session_state.ANALYSIS_VARIABLES_TEXTS["COLAB_POPUP_NEXT_STEPS"], key="loc_an_next_steps")
            st.session_state.ANALYSIS_VARIABLES_TEXTS["COLAB_POPUP_STEP1_PART1"] = st.text_input(_("an_colab_step1_part1_label", T_UI), st.session_state.ANALYSIS_VARIABLES_TEXTS["COLAB_POPUP_STEP1_PART1"], key="loc_an_s1p1")
            st.session_state.ANALYSIS_VARIABLES_TEXTS["COLAB_POPUP_STEP1_LINK_TEXT"] = st.text_input(_("an_colab_step1_link_text_label", T_UI), st.session_state.ANALYSIS_VARIABLES_TEXTS["COLAB_POPUP_STEP1_LINK_TEXT"], key="loc_an_s1link")
            st.session_state.ANALYSIS_VARIABLES_TEXTS["COLAB_POPUP_STEP1_SMALL_TEXT"] = st.text_area(_("an_colab_step1_small_text_label", T_UI), st.session_state.ANALYSIS_VARIABLES_TEXTS["COLAB_POPUP_STEP1_SMALL_TEXT"], key="loc_an_s1small")
            st.session_state.ANALYSIS_VARIABLES_TEXTS["COLAB_POPUP_STEP2"] = st.text_input(_("an_colab_step2_label", T_UI), st.session_state.ANALYSIS_VARIABLES_TEXTS["COLAB_POPUP_STEP2"], key="loc_an_s2")
            st.session_state.ANALYSIS_VARIABLES_TEXTS["COLAB_POPUP_STEP3"] = st.text_input(_("an_colab_step3_label", T_UI), st.session_state.ANALYSIS_VARIABLES_TEXTS["COLAB_POPUP_STEP3"], key="loc_an_s3")

        with st.expander(_("colors_editor_label", T_UI)):
            st.caption(_("data_editor_info_dict", T_UI))
            if isinstance(st.session_state.MODERATOR_VARIABLES["COLORS"], dict):
                st.session_state.MODERATOR_VARIABLES["COLORS"] = st.data_editor(
                    st.session_state.MODERATOR_VARIABLES["COLORS"], key="fc_colors_editor"
                )
            else: st.warning(_("warning_not_dict", T_UI, variable_name="MODERATOR_VARIABLES.COLORS"))

        with st.expander(_("data_prep_options_header", T_UI)):
            st.session_state.MODERATOR_VARIABLES["DATA_PREP"]["SheetName"] = st.text_input(
                _("prepdata_sheet_label", T_UI), st.session_state.MODERATOR_VARIABLES["DATA_PREP"]["SheetName"], key="qs_prepdata_sheet"
            )
            st.session_state.colabGitHubUrl = st.text_input(
                _("colab_url_label", T_UI), st.session_state.colabGitHubUrl, key="qs_colab_url"
            )
            dp_vars = st.session_state.MODERATOR_VARIABLES["DATA_PREP"]
            lang_options_map = { "English (no translation)": "", "Auto-detect (Google Translate)": "auto",
                                "Danish": "da", "Dutch": "nl", "French": "fr", "German": "de",
                                "Italian": "it", "Portuguese": "pt", "Spanish": "es", "Swedish": "sv"}
            current_s_lang_val = dp_vars["SessionLanguage"]
            current_s_disp_lang = next((k for k, v in lang_options_map.items() if v == current_s_lang_val), list(lang_options_map.keys())[0])
            selected_s_disp_lang = st.selectbox(
                _("session_language_label", T_UI), options=list(lang_options_map.keys()),
                index=list(lang_options_map.keys()).index(current_s_disp_lang), key="fc_session_lang"
            )
            dp_vars["SessionLanguage"] = lang_options_map[selected_s_disp_lang]

            target_lang_options = {"English": "en", "French": "fr", "German": "de", "Spanish": "es"}
            current_t_lang_val = dp_vars["TranslatedLanguage"]
            current_t_disp_lang = next((k for k, v in target_lang_options.items() if v == current_t_lang_val), "English")
            selected_t_disp_lang = st.selectbox(
                _("translated_language_label", T_UI), options=list(target_lang_options.keys()),
                index=list(target_lang_options.keys()).index(current_t_disp_lang), key="fc_target_lang"
            )
            dp_vars["TranslatedLanguage"] = target_lang_options[selected_t_disp_lang]

            dp_vars["IdeaRawColumn"] = st.text_input(_("idea_raw_column_label", T_UI), dp_vars["IdeaRawColumn"], key="fc_raw_col")
            dp_vars["TranslateColumn"] = st.text_input(_("translate_column_label", T_UI), dp_vars["TranslateColumn"], key="fc_trans_col")
            dp_vars["ManualCategorization"] = st.text_input(_("manual_categorization_label", T_UI), dp_vars["ManualCategorization"], key="fc_manual_cat_col")


    # --- Generate and Download GS ---
    st.divider()
    if st.button(_("generate_js_button", T_UI), type="primary", key="download_button_main"):
        generated_js = generate_js_from_state(T_UI)
        if generated_js:

            st.download_button(
                label=_("download_js_button", T_UI),
                data=generated_js.encode("utf-8").decode("utf-8"),
                file_name="configured_brainwriting_script.gs",
                mime="application/javascript",
                key="download_action_button"
            )
            st.success(_("js_generated_success", T_UI))
            st.code(generated_js, language="javascript")

if __name__ == "__main__":
    # Need to import pandas for st.data_editor with DataFrame
    try:
        import pandas as pd
    except ImportError:
        st.error("Pandas library is not installed. Please install it: pip install pandas")
        st.stop()
    main()
