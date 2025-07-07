# Customizable Structured Electronic Brainwriting Session: Utilizing Google Sheets with Google Apps Script

[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)
[![Python Version](https://img.shields.io/badge/python-3.8+-blue.svg)](https://www.python.org/downloads/)

# Structured Electronic Brainwriting Session Configuration Generator

A powerful Streamlit application that generates JavaScript configuration files for customizable structured electronic brainwriting (CSEB) and multilingual CSEB sessions. This tool simplifies the customization process for facilitators running structured electronic brainstorming sessions, with support for multiple languages and advanced configuration options.

## Features

- **Multilingual and Monolingual Sessions**:
  - **Multiple Language Support**: Participants contribute the session in their mothertongue
  - **Simpler Monolingual Sessions**: Generate essential configurations in minutes with minimal inputs

- **Three Configuration Modes**:
  - **Quick Setup**: Generate essential configurations in minutes with minimal inputs
  - **Variable Localization**: Participant visible text to be configured. Translate configurations to other languages here for localization. (English, Turkish, German, Spanish, and French languages are currently translated)
  - **Advanced Customization**: Access all variables for complete control over your sessions

- **Smart Defaults**:
  - Auto-generates participant names, round labels, and idea placeholders
  - Intelligent formatting of sheet names and variable names
  - Pre-configured color schemes and UI labels

- **User-Friendly Interface**:
  - Real-time validation and preview of configurations
  - Automatic formatting of user inputs to prevent errors
  - Interactive elements for better usability

- **Output Options**:
  - Download complete  App Script configuration file
  - Copy code directly to clipboard
  - Preview generated code before finalizing


## Before starting
### Research the Focus Area:

Before kicking off the brainwriting session, take some time to dive deep into your research question or focus area. It's really important to have a solid grasp of the topic and what you hope to achieve in the session.
Put forward a clear and straightforward problem statement. This will help participants get engaged and stay focused on coming up with new ideas.

### Prepare for Implementation:

Familiarize yourself with the Google Apps Script, ensure you understand the script provided, and make any necessary changes to fit your requirements.

### Test the Setup:

Make sure to tackle any bugs or limitations you find before you work with your actual participants. Conduct a trial run with a small group or colleagues to identify any potential issues with the script, permissions, or workflow.

### Ensure Participant Readiness:

Let participants know what to expect from the session ahead of time.
Make sure everyone has a stable internet connection and can access Google Sheets without any issues.

## Technical Details

This application uses:
- Streamlit for the user interface
- JavaScript templates with placeholder replacement
- Pre-defined translation dictionaries
- Input validation and sanitization

## Integration

The generated App Script configuration files are designed to work with electronic brainwriting systems in Google Sheets. Simply copy the generated file and paste it to Google Apps Script, or download the code for further configuration.

## Usage

1. **Choose Multilingual or Monolingual Variation**:
   Multilingual session feature is powered by GOOGLETRANSLATE function in the Google Sheets.

2. **Select Configuration Mode**:
   Choose between Quick Setup, Variable Localization, or Advanced Customization based on your needs.

3. **Enter Session Details**:
   - Define session focus/topic
   - Set number of participants (and optionally customize names)
   - Configure rounds, ideas per round, and time limits

4. **Customize as Needed**:
   - Add a logo for the landing page
   - Select language for session interface
   - Adjust colors and labels (in full customization mode)

5. **Generate and Download**:
   - Preview the generated  App Script configuration
   - Download the file or copy the code
   - Integrate with your electronic brainwriting system

6. **Analyze and Save**:
   - Use preconfigured Colab Notebook for advanced analysis in Python
   - Data load and save back to Google Sheets functions are provided

## Example

Here's a quick example of how to generate a basic configuration:

1. Select "Quick Setup (English)"
2. Enter session focus: "Improving Remote Team Collaboration"
3. Set 7 participants, 5 rounds, 2 ideas per round, 3 minutes per round
4. Click "Generate Script"
5. Copy the App Script file
6. Open a new blank Google Sheets
7. Navigate to Apps Script
8. Paste the generated code
9. "Save to Drive" and "Run"
10. Follow further steps in our protocol: https://dx.doi.org/10.17504/protocols.io.x54v9r3xzv3e/v1

The generated configuration will handle all the naming, formatting, and calculations needed for your session.

## Language Support

Multilingual sessions support all GOOGLETRANSLATE languages:
{
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

The CSEB Session Configuration Generator application currently supports the following languages:
- English (Default)
- Turkish
- German
- Spanish
- French

When using language localization, all interface elements, prompts, and generated text will be available in the selected language.

## Contributing

Contributions are welcome! Please feel free to submit a Pull Request.
Translations and localizations are also welcome. Right after the initial compatibility check, the application will be updated. Please use /locales/en.json file as the main language to be translated.

## License

This project is licensed under the MIT License - see the LICENSE file for details.

## Acknowledgements

- Built with [Streamlit](https://streamlit.io/)
- Inspired by collaborative brainstorming techniques
- Special thanks to contributors of the brainwriting methodology

## Step-by-step implementation is in protocols.io:

https://dx.doi.org/10.17504/protocols.io.x54v9r3xzv3e/v1

## Abstract of the Protocol
This protocol outlines the process and structure to use Google Sheets along with Google Apps Script to set up and manage an online 6-3-5 brainwriting session, which enables remote collaborative sessions for idea generation. This technique creates a structured online platform by automating traditional brainwriting sessions, which involve six participants each coming up with three ideas per round and passing their paper round-robin to the next participant, building new ideas upon previous contributions over the course of six rounds.

Google Sheets functions as a versatile platform for real-time input collection and documentation while enabling participants to effortlessly browse on the spreadsheet and contribute to the idea development process. With this approach, moderators can efficiently guide and document brainstorming sessions, whether they’re in-person or virtual. The online 6-3-5 brainwriting system is designed to automatically track progress, organize participant contributions, and gather a total of 108 ideas—thanks to 6 participants collaborating through 6 rounds, with each of them sharing 3 ideas per round—all neatly organized in a structured worksheet.

## Reference

Çetin, Mahmut Berkan & Gündüz, Selim (2025). Protocol for Conducting Structured Electronic Brainwriting Session: Utilizing Google Sheets with Google Apps Script. protocols.io, https://dx.doi.org/10.17504/protocols.io.x54v9r3xzv3e/v1
