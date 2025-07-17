# Customizable Structured Electronic Brainwriting Session: Utilizing Google Sheets with Google Apps Script

[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)
[![Python Version](https://img.shields.io/badge/python-3.8+-blue.svg)](https://www.python.org/downloads/)
[![Streamlit App](https://img.shields.io/badge/Streamlit-App-red?logo=streamlit)]([https://streamlit.app.link/your-app](https://cseb-configuration-generator.streamlit.app/))

## Structured Electronic Brainwriting Session Configuration Generator

A powerful Streamlit application that generates JavaScript configuration files for customizable structured electronic brainwriting (CSEB) and multilingual CSEB sessions. This tool simplifies the customization process for facilitators running structured electronic brainstorming sessions, with support for multiple languages and advanced configuration options.

### Features

The CSEB Session Configuration Generator provides a comprehensive set of features to tailor your brainwriting sessions:

*   **Multilingual and Monolingual Sessions**:
    *   **Multiple Language Support**: Facilitates sessions where participants can contribute in their mother tongue.
    *   **Simpler Monolingual Sessions**: Enables the generation of essential configurations in minutes with minimal inputs for single-language sessions.

*   **Three Configuration Modes**:
    *   **Quick Setup**: Allows users to generate essential configurations swiftly with minimal inputs.
    *   **Variable Localization**: Provides options to configure participant-visible text and translate configurations into other languages for interface localization. English, Turkish, German, Spanish, and French are currently translated for the application interface.
    *   **Advanced Customization**: Offers access to all variables, providing complete control over session parameters.

*   **Smart Defaults**:
    *   Automatically generates participant names, round labels, and idea placeholders.
    *   Applies intelligent formatting for sheet names and variable names.
    *   Includes pre-configured color schemes and UI labels.

*   **User-Friendly Interface**:
    *   Features real-time validation and preview of configurations.
    *   Provides automatic formatting of user inputs to prevent errors.
    *   Incorporates interactive elements for enhanced usability.

*   **Output Options**:
    *   Allows downloading the complete App Script configuration file.
    *   Enables direct copying of the generated code to the clipboard.
    *   Offers a preview of the generated code before finalization.

### How to Use the Generator (Usage Guide)

To generate and implement your customized electronic brainwriting session script:

1.  **Choose Multilingual or Monolingual Variation**: Select the desired session type. The multilingual session feature is powered by the `GOOGLETRANSLATE` function in Google Sheets.
2.  **Select Configuration Mode**: Choose between Quick Setup, Variable Localization, or Advanced Customization based on your specific needs.
3.  **Enter Session Details**: Define the session focus/topic, set the number of participants (with an option to customize names), and configure the number of rounds, ideas per round, and time limits.
4.  **Customize as Needed**: Further customize by adding a logo for the landing page, selecting the language for the session interface, and adjusting colors and labels in full customization mode.
5.  **Generate and Download**: Preview the generated App Script configuration, then download the file or copy the code.
6.  **Integrate with Google Sheets**:
    *   Open a new blank Google Sheet.
    *   Navigate to Apps Script.
    *   Paste the generated code.
    *   "Save to Drive" and "Run".
7.  **Analyze and Save**: Utilize the preconfigured Colab Notebook for advanced analysis in Python. Data load and save functions back to Google Sheets are provided to streamline your analysis workflow.

The generated configuration will automatically handle all the necessary naming, formatting, and calculations for your session.

### Technical Details

This application is built using:
*   **Streamlit**: For the user interface.
*   **Apps Script templates**: With placeholder replacement for generating configuration files.
*   **Pre-defined translation dictionaries**: For internationalization.
*   **Input validation and sanitization**: To ensure robust and error-free configuration.
*   The application's backend is primarily Python [19, 20].

### Integration

The generated App Script configuration files are specifically designed to work seamlessly with electronic brainwriting systems within Google Sheets. Users simply copy the generated file and paste it into Google Apps Script, or download the code for further configuration.

### Language Support

*   **Multilingual sessions** are designed to support all languages available via the `GOOGLETRANSLATE` function. This allows participants to contribute in a wide array of mother tongues.
*   The **CSEB Session Configuration Generator application's interface** itself currently supports English (Default),Turkish, German, Spanish, French languages for localization. When using language localization for the application, all interface elements, prompts, and generated text will be available in the selected language.

### Contributing

Contributions are welcome! Please feel free to submit a Pull Request. Translations and localizations are also welcome. Right after the initial compatibility check, the application will be updated. Please use `/locales/en.json` file as the main language to be translated.

### License

This project is licensed under the MIT License - see the `LICENSE` file for details.

### Acknowledgements

- Built with [Streamlit](https://streamlit.io/)
- Inspired by collaborative brainstorming techniques
- Special thanks to contributors of the brainwriting methodology

### Associated Protocol and Methodology

**For a detailed academic protocol and step-by-step implementation guide on how to conduct a Structured Electronic Brainwriting Session utilizing Google Sheets with Google Apps Script, please refer to the following publication on protocols.io**:

Çetin, Mahmut Berkan & Gündüz, Selim (2025). Protocol for Conducting Structured Electronic Brainwriting Session: Utilizing Google Sheets with Google Apps Script. protocols.io, https://dx.doi.org/10.17504/protocols.io.x54v9r3xzv3e/v1

This protocol outlines the process and structure for setting up and managing an online 6-3-5 brainwriting session, enabling remote collaborative idea generation. The technique automates traditional brainwriting, where six participants each generate three ideas per round and pass their contributions to the next person, building upon previous ideas over six rounds. Google Sheets serves as a versatile platform for real-time input collection and documentation, allowing participants to easily browse and contribute to idea development. This approach empowers moderators to efficiently guide and document brainstorming sessions, both in-person and virtually. The online 6-3-5 brainwriting system automatically tracks progress, organizes participant contributions, and can gather a total of 108 ideas (from 6 participants collaborating through 6 rounds, with 3 ideas per round), all neatly organized in a structured worksheet.
