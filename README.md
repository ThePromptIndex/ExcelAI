# ExcelAI
ExcelGPT Assistant ExcelGPT Assistant is a VBA-based tool that integrates with the OpenAI GPT-4 API to provide intelligent responses directly within the Excel environment. This assistant can help users with a variety of Excel tasks by providing accurate and context-aware responses.

**Features**
Seamless Integration: Communicates with the OpenAI GPT-4 API to fetch responses.
Customizable Prompts: Allows users to input queries and specify a word limit for responses.
Error Handling: Includes robust error handling for API requests and response parsing.
Excel-Friendly Output: Adjusts responses to fit well within Excel, including handling formula responses appropriately.

**Usage**
Set Up: Ensure you have a valid OpenAI API key and update the apiKey variable in the code.
Run the Function: Use the AskAI function to send a query to the GPT-4 API and receive a response.
Customizable Queries: You can specify a maximum word count for the response by passing an optional parameter.
Code Overview
AskAI: Main function to send a query and retrieve the response.
ParseContent: Extracts the relevant content from the API response.
ParseError: Extracts error messages from the API response for debugging purposes.
**Installation**
Open your Excel workbook.
Press ALT + F11 to open the VBA editor.
Insert a new module and copy-paste the VBA code from this repository.
Update the apiKey variable with your OpenAI API key.
Save as a macro enabled workbook
Contributing
Feel free to submit issues or pull requests if you have suggestions or improvements. Contributions are always welcome!
