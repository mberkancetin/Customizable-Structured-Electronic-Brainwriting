# Customizable Structured Electronic Brainwriting Session: Utilizing Google Sheets with Google Apps Script

[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)
[![Python Version](https://img.shields.io/badge/python-3.8+-blue.svg)](https://www.python.org/downloads/)

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

# Structured Electronic Brainwriting Session Configuration Generator

A powerful Streamlit application that generates JavaScript configuration files for customizable structured electronic brainwriting (CSEB) sessions. This tool simplifies the customization process for facilitators running structured electronic brainstorming sessions, with support for multiple languages and advanced configuration options.

## Features

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

## Usage

1. **Select Configuration Mode**:
   Choose between Quick Setup, Variable Localization, or Advanced Customization based on your needs.

2. **Enter Session Details**:
   - Define session focus/topic
   - Set number of participants (and optionally customize names)
   - Configure rounds, ideas per round, and time limits

3. **Customize as Needed**:
   - Add a logo for the landing page
   - Select language for session interface
   - Adjust colors and labels (in full customization mode)

4. **Generate and Download**:
   - Preview the generated  App Script configuration
   - Download the file or copy the code
   - Integrate with your electronic brainwriting system

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

The application currently supports the following languages:
- English (Default)
- Turkish
- German
- Spanish
- French

When using language localization, all interface elements, prompts, and generated text will be available in the selected language.

## Technical Details

This application uses:
- Streamlit for the user interface
- JavaScript templates with placeholder replacement
- Pre-defined translation dictionaries
- Input validation and sanitization

## Integration

The generated App Script configuration files are designed to work with electronic brainwriting systems in Google Sheets. Simply copy the generated file and paste it to Google Apps Script, or download the code for further configuration.

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

https://protocols.io/view/online-6-3-5-brainwriting-session-utilizing-google-dwwn7fde

## Abstract
This protocol outlines the process and structure to use Google Sheets along with Google Apps Script to set up and manage an online 6-3-5 brainwriting session, which enables remote collaborative sessions for idea generation. This technique creates a structured online platform by automating traditional brainwriting sessions, which involve six participants each coming up with three ideas per round and passing their paper round-robin to the next participant, building new ideas upon previous contributions over the course of six rounds. 

Google Sheets functions as a versatile platform for real-time input collection and documentation while enabling participants to effortlessly browse on the spreadsheet and contribute to the idea development process. With this approach, moderators can efficiently guide and document brainstorming sessions, whether they’re in-person or virtual. The online 6-3-5 brainwriting system is designed to automatically track progress, organize participant contributions, and gather a total of 108 ideas—thanks to 6 participants collaborating through 6 rounds, with each of them sharing 3 ideas per round—all neatly organized in a structured worksheet.

## Reference

Çetin, Mahmut Berkan & Gündüz, Selim (2025). Protocol for Conducting Structured Electronic Brainwriting Session: Utilizing Google Sheets with Google Apps Script. protocols.io, https://dx.doi.org/10.17504/protocols.io.x54v9r3xzv3e/v1
