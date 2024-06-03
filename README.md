# ExcelAI Assistant

ExcelAI Assistant is a VBA-based tool that integrates with the OpenAI GPT-4 API to provide intelligent responses directly within the Excel environment. This assistant can help users with a variety of Excel tasks by providing accurate and context-aware responses.

![Untitled design (4)](https://github.com/ThePromptIndex/ExcelAI/assets/144904800/8159ff51-52b8-4001-9cfe-f111163b9ced)

If you enjoyed this sort of stuff then [sign up to our newsletter](https://www.thepromptindex.com/newsletter.html).

**Features**

- **Seamless Integration**: Communicates with the OpenAI GPT-4 API to fetch responses.
- **Customizable Prompts**: Allows users to input queries and specify a word limit for responses.
- **Error Handling**: Includes robust error handling for API requests and response parsing.
- **Excel-Friendly Output**: Adjusts responses to fit well within Excel, including handling formula responses appropriately.

**Usage**

1. **Set Up**: Ensure you have a valid OpenAI API key and update the `apiKey` variable in the code.
2. **Run the Function**: Use the `AskAI` function to send a query to the GPT-4 API and receive a response.
3. **Customizable Queries**: You can specify a maximum word count for the response by passing an optional parameter.

**Code Overview**

- `AskAI`: Main function to send a query and retrieve the response.
- `ParseContent`: Extracts the relevant content from the API response.
- `ParseError`: Extracts error messages from the API response for debugging purposes.

**Installation**

1. Open your Excel workbook.
2. Press `ALT + F11` to open the VBA editor.
3. Insert a new module and copy-paste the VBA code from this repository.
4. Update the `apiKey` variable with your OpenAI API key.
5. Save as a macro-enabled workbook.

**Contributing**

Feel free to submit issues or pull requests if you have suggestions or improvements. Contributions are always welcome!
