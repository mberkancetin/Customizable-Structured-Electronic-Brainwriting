# SPDX-FileCopyrightText: 2025 Mahmut Berkan Çetin <m.berkancetin@gmail.com> & Selim Gündüz <sgunduz@atu.edu.tr>
#
# SPDX-License-Identifier: MIT

# KEYS for default text that needs translation
DEFAULT_TEXT_KEYS = {
    "SESSION_FOCUS": "default_session_focus",
    "LANDING_SHEET": "default_landing_sheet_name", # Or keep as "Welcome" if sheet names aren't translated
    "PARTICIPANT_PREFIX": "default_participant_prefix", # For generating Participant1, Participant2 etc.
    "PARTICIPANT_SHEET_LANGUAGE": ["default_session_language", "default_session_language", "default_session_language", "default_session_language", "default_session_language", "default_session_language", "default_mod_menu_reset_session_setup"],
    "PARTICIPANT_SHEET_FOCUS": ["default_session_focus", "default_session_focus", "default_session_focus", "default_session_focus", "default_session_focus", "default_session_focus"],
    "MODERATOR_SHEET": "default_moderator_sheet_name",
    "SESSION_LANGUAGE": "default_session_language",
    "IDEA_SWAP_ALGORITHM": "default_idea_swap_algorithm",
    "TIME_LEFT": "default_time_left",
    "MINS_LEFT": "default_mins_left",
    "ONE_MIN_LEFT": "default_one_min_left",
    "TIME_IS_UP": "default_time_is_up",
    "FOCUS_ITEM_1": "default_focus_item_1", # "Focus"
    "FOCUS_ITEM_2": "default_focus_item_2", # "Session Started"
    "FOCUS_ITEM_3": "default_focus_item_3", # "Session Completed"
    "ROUND_PREFIX": "default_round_prefix", # For "Round 1", "Round 2"
    "IDEA_PREFIX": "default_idea_prefix",   # For "Idea 1", "Idea 2"
    "CHECK_IDEAS": "default_check_ideas",
    "ROUND_CHANGE_MSG1": "default_round_change_msg1",
    "ROUND_CHANGE_MSG2": "default_round_change_msg2",
    "SESSION_COMPLETE": "default_session_complete",
    "STARTING": "default_glob_starting",
    "STOPPED": "default_glob_stopped",

    # Moderator Variables Texts
    "MOD_TOOLS": "default_mod_menu_tools",
    "MOD_LANDING_PAGE": "default_mod_menu_landing_page",
    "MOD_SESSION_ELEMENTS": "default_mod_menu_session_elements",
    "MOD_START": "default_mod_menu_start",
    "MOD_SUBMIT_NEXT": "default_mod_menu_submit_next",
    "MOD_PREPARE_DATA": "default_mod_menu_prepare_data",
    "MOD_COLAB_ENV": "default_mod_menu_colab_env",
    "MOD_RESET_SETUP": "default_mod_menu_reset_session_setup",
    "MOD_SESSION_START_TEMPLATE": "default_mod_session_start_template", # e.g., "Please submit {ideasCount} ideas..."
    "MOD_ROUND_END_PHRASE": "default_mod_round_end_phrase",
    "MOD_CURRENT_ROUND": "default_mod_current_round",
    "MOD_DATA_PREP_SHEET_NAME": "default_mod_data_prep_sheet_name", # Or keep as "PrepData"
    "MOD_DATA_PREP_IDEA_RAW_COL": "default_mod_data_prep_idea_raw_col", # "RawIdea"
    "MOD_DATA_PREP_TRANSLATE_COL": "default_mod_data_prep_translate_col", # "Translation"
    "MOD_DATA_PREP_IDEAID_COL": "default_mod_data_prep_ideaid_col",
    "MOD_DATA_PREP_IDEATIMESTAMP_COL": "default_mod_data_prep_ideatimestamp_col",
    "MOD_DATA_PREP_ROUNDTIMESTAMP_COL": "default_mod_data_prep_roundtimestamp_col",
    "MOD_DATA_PREP_MANUAL_CAT": "default_mod_data_prep_manual_cat", # "ManualCategories"
    "MOD_START_TIMER": "default_mod_start_timer",

    # Landing Page Texts (content of the variables, not labels)
    "LP_GREETING_MESSAGE_TEMPLATE": "default_lp_greeting_message_template", # The big template string
    "LP_SESSION_TITLE_SUFFIX": "default_lp_session_title_suffix",
    "LP_PARTICIPANTS_PREFIX": "default_lp_participants_prefix",
    "LP_ROUNDS_PREFIX": "default_lp_rounds_prefix",
    "LP_IDEAS_PREFIX": "default_lp_ideas_prefix",
    "LP_TIME_PREFIX": "default_lp_time_prefix",
    "LP_TOTALS_INFIX": "default_lp_totals_infix",

    # Analysis Variables Texts
    "AN_POPUP_TITLE": "default_an_popup_title",
    "AN_COLAB_NEXT_STEPS": "default_an_colab_next_steps",
    "AN_COLAB_STEP1_PART1": "default_an_colab_step1_part1",
    "AN_COLAB_STEP1_LINK_TEXT": "default_an_colab_step1_link_text",
    "AN_COLAB_STEP1_SMALL_TEXT": "default_an_colab_step1_small_text",
    "AN_COLAB_STEP2": "default_an_colab_step2",
    "AN_COLAB_STEP3": "default_an_colab_step3",

    "ALERT_LANDING_PAGE_HEADER": "landing_page_alert_header",
    "ALERT_LANDING_PAGE_TEXT": "landing_page_alert_text",
    "ALERT_CREATE_SESSION_HEADER": "create_session_alert_header",
    "ALERT_CREATE_SESSION_TEXT": "create_session_alert_text",
    "ALERT_RESET_SESSION_HEADER": "reset_session_flags_header",
    "ALERT_RESET_SESSION_TEXT": "reset_session_flags_text",
}

# Non-translatable defaults or structural defaults
DEFAULT_GLOBAL_VARIABLES_STRUCTURE = {
    "SESSION_FOCUS": DEFAULT_TEXT_KEYS["SESSION_FOCUS"], # Placeholder for text
    "LANDING_SHEET": DEFAULT_TEXT_KEYS["LANDING_SHEET"],
    "PARTICIPANT": ["Participant1_key", "Participant2_key"], # Number defined by num_participants, names generated
    "PARTICIPANT_SHEET_LANGUAGE": DEFAULT_TEXT_KEYS["PARTICIPANT_SHEET_LANGUAGE"],
    "PARTICIPANT_SHEET_FOCUS": DEFAULT_TEXT_KEYS["PARTICIPANT_SHEET_FOCUS"],
    "MODERATOR_SHEET": DEFAULT_TEXT_KEYS["MODERATOR_SHEET"],
    "SESSION_LANGUAGE": DEFAULT_TEXT_KEYS["SESSION_LANGUAGE"],
    "TIME_LEFT": DEFAULT_TEXT_KEYS["TIME_LEFT"],
    "MINS_LEFT": DEFAULT_TEXT_KEYS["MINS_LEFT"],
    "ONE_MIN_LEFT": DEFAULT_TEXT_KEYS["ONE_MIN_LEFT"],
    "TIME_IS_UP": DEFAULT_TEXT_KEYS["TIME_IS_UP"],
    "MINUTES": 5, # Integer, not translated
    "FOCUS": [DEFAULT_TEXT_KEYS["FOCUS_ITEM_1"], DEFAULT_TEXT_KEYS["FOCUS_ITEM_2"], DEFAULT_TEXT_KEYS["FOCUS_ITEM_3"]],
    "ROUNDS": ["Round1_key", "Round2_key"], # Number defined by num_rounds, names generated
    "IDEAS": ["Idea1_key", "Idea2_key"],   # Number defined by num_ideas, names generated
    "CHECK_IDEAS": DEFAULT_TEXT_KEYS["CHECK_IDEAS"],
    "ROUND_CHANGE": [DEFAULT_TEXT_KEYS["ROUND_CHANGE_MSG1"], DEFAULT_TEXT_KEYS["ROUND_CHANGE_MSG2"]],
    "SESSION_COMPLETE": DEFAULT_TEXT_KEYS["SESSION_COMPLETE"],
    "STARTING": DEFAULT_TEXT_KEYS["STARTING"],
    "STOPPED": DEFAULT_TEXT_KEYS["STOPPED"],
}

DEFAULT_MODERATOR_VARIABLES_STRUCTURE = {
    "MENU": { # Menu item text will be translated
        "Tools": DEFAULT_TEXT_KEYS["MOD_TOOLS"],
        "LandingPage": DEFAULT_TEXT_KEYS["MOD_LANDING_PAGE"],
        "SessionElements": DEFAULT_TEXT_KEYS["MOD_SESSION_ELEMENTS"],
        "Start": DEFAULT_TEXT_KEYS["MOD_START"],
        "SubmitNext": DEFAULT_TEXT_KEYS["MOD_SUBMIT_NEXT"],
        "PrepareData": DEFAULT_TEXT_KEYS["MOD_PREPARE_DATA"],
        "ColabEnvironment": DEFAULT_TEXT_KEYS["MOD_COLAB_ENV"],
        "ResetSessionSetup": DEFAULT_TEXT_KEYS["MOD_RESET_SETUP"],
    },
    "SESSION_START_TEMPLATE": DEFAULT_TEXT_KEYS["MOD_SESSION_START_TEMPLATE"],
    "ROUND_END_PHRASE": DEFAULT_TEXT_KEYS["MOD_ROUND_END_PHRASE"],
    "CURRENT_ROUND": DEFAULT_TEXT_KEYS["MOD_CURRENT_ROUND"],
    "DATA_PREP": { # Some are sheet/column names (might not translate), some are config values
        "SheetName": DEFAULT_TEXT_KEYS["MOD_DATA_PREP_SHEET_NAME"],
        "SessionLanguage": "auto", # Config value
        "IdeaRawColumn": DEFAULT_TEXT_KEYS["MOD_DATA_PREP_IDEA_RAW_COL"],
        "TranslatedLanguage": "en", # Config value
        "TranslateColumn": DEFAULT_TEXT_KEYS["MOD_DATA_PREP_TRANSLATE_COL"],
        "IdeaID": DEFAULT_TEXT_KEYS["MOD_DATA_PREP_IDEAID_COL"],
        "IdeaTimestamp": DEFAULT_TEXT_KEYS["MOD_DATA_PREP_IDEATIMESTAMP_COL"],
        "IdeaRoundStartTimestamp": DEFAULT_TEXT_KEYS["MOD_DATA_PREP_ROUNDTIMESTAMP_COL"],
        "ManualCategorization": DEFAULT_TEXT_KEYS["MOD_DATA_PREP_MANUAL_CAT"]
    },
    "COLORS": { # Not translated
        "LightSteelBlue": "LightSteelBlue", "LightBlue": "#cfe2f3", "Gold": "#ffd700",
        "LightGreen": "#d9ead3", "Green": "#90ee90", "Yellow": "yellow",
        "Grey": "grey", "LightGrey": "#efefef"
    },
    "START_TIMER": DEFAULT_TEXT_KEYS["MOD_START_TIMER"]
}

DEFAULT_LANDINGPAGE_VARIABLES_TEXT_KEYS = { # Keys for the text parts
    "GREETING_MESSAGE_TEMPLATE": DEFAULT_TEXT_KEYS["LP_GREETING_MESSAGE_TEMPLATE"],
    "SESSION_TITLE_SUFFIX": DEFAULT_TEXT_KEYS["LP_SESSION_TITLE_SUFFIX"],
    "PARTICIPANTS_PREFIX": DEFAULT_TEXT_KEYS["LP_PARTICIPANTS_PREFIX"],
    "ROUNDS_PREFIX": DEFAULT_TEXT_KEYS["LP_ROUNDS_PREFIX"],
    "IDEAS_PREFIX": DEFAULT_TEXT_KEYS["LP_IDEAS_PREFIX"],
    "TIME_PREFIX": DEFAULT_TEXT_KEYS["LP_TIME_PREFIX"],
    "TOTALS_TEXT_INFIX": DEFAULT_TEXT_KEYS["LP_TOTALS_INFIX"],
}

DEFAULT_ANALYSIS_VARIABLES_TEXT_KEYS = {
    "POPUP_TITLE": DEFAULT_TEXT_KEYS["AN_POPUP_TITLE"],
    "COLAB_POPUP_NEXT_STEPS": DEFAULT_TEXT_KEYS["AN_COLAB_NEXT_STEPS"],
    "COLAB_POPUP_STEP1_PART1": DEFAULT_TEXT_KEYS["AN_COLAB_STEP1_PART1"],
    "COLAB_POPUP_STEP1_LINK_TEXT": DEFAULT_TEXT_KEYS["AN_COLAB_STEP1_LINK_TEXT"],
    "COLAB_POPUP_STEP1_SMALL_TEXT": DEFAULT_TEXT_KEYS["AN_COLAB_STEP1_SMALL_TEXT"],
    "COLAB_POPUP_STEP2": DEFAULT_TEXT_KEYS["AN_COLAB_STEP2"],
    "COLAB_POPUP_STEP3": DEFAULT_TEXT_KEYS["AN_COLAB_STEP3"],
}

DEFAULT_ALERT_LANDING_PAGE_HEADER = DEFAULT_TEXT_KEYS["ALERT_LANDING_PAGE_HEADER"]
DEFAULT_ALERT_LANDING_PAGE_TEXT = DEFAULT_TEXT_KEYS["ALERT_LANDING_PAGE_TEXT"]
DEFAULT_ALERT_CREATE_SESSION_HEADER = DEFAULT_TEXT_KEYS["ALERT_CREATE_SESSION_HEADER"]
DEFAULT_ALERT_CREATE_SESSION_TEXT = DEFAULT_TEXT_KEYS["ALERT_CREATE_SESSION_TEXT"]
DEFAULT_ALERT_RESET_SESSION_HEADER = DEFAULT_TEXT_KEYS["ALERT_RESET_SESSION_HEADER"]
DEFAULT_ALERT_RESET_SESSION_TEXT = DEFAULT_TEXT_KEYS["ALERT_RESET_SESSION_TEXT"]

# Non-translatable defaults
DEFAULT_IMAGE_URL = "https://upload.wikimedia.org/wikipedia/commons/f/f3/Adana_Alparslan_T%C3%BCrke%C5%9F_%C3%9Cniversitesi_logo.png"
DEFAULT_COLAB_GITHUB_URL = "https://colab.research.google.com/github/mberkancetin/Customizable-Structured-Electronic-Brainwriting/blob/main/DataAnalysis.ipynb"

# Number of default participants, rounds, ideas
DEFAULT_NUM_PARTICIPANTS = 6
DEFAULT_NUM_ROUNDS = 6
DEFAULT_NUM_IDEAS = 3
