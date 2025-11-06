# ChatGPT for Excel
ChatGPT for Excel using VBA, bringing the power of AI to any version of Excel.

## Usage
Formula requires a `prompt`, `model`, `reasoning-level`, and `api-key`:
```excel
=ChatGPT(prompt, model, reasoning-level, api-key)
```

### Basic example
Paste the following code in a cell in Excel to get a response:
```excel
=ChatGPT("How high is the Eifel Tower?", "gpt-5-nano", "low", "YOUR-OPENAI-API-KEY")
```

The result will be similar to:
>The Eiffel Tower in Paris is about 330 meters (1,083 feet) tall, including its antennas. The original height when it was completed in 1889 was 300 meters, but antennas added later increased it.

*Note: OpenAI can take a few seconds (per query) to return a response. To get a different response, tune the prompt.*

### Example with cell reference
The prompt, model, reasoning, and api-key can also be located in an Excel cell, e.g. cell `A1` holds the prompt, cell `B1` the model, `C1` the reasoning level, and `D1` the OpenAI api-key. To get a response, paste the following in cell `E1`:
```excel
=ChatGPT(A1, B1, C1, D1)
```



### Model selection (and price per 1M tokens)
Available models and their associated costs can be found on: [OpenAi pricing](https://platform.openai.com/docs/pricing)

### Reasoning level
It is possible to set the reasoning level for the model to `low`, `medium`, or `high`.

For more information see: [OpenAi reasoning](https://platform.openai.com/docs/guides/reasoning)

### Word of caution
It is **strongly** advised that after a response has been generated to **hard paste** the value. Failing to do so will re-query the OpenAI-api on *each and every*  change in the Excel, both temporarily freezing Excel until the response is returned and costing money for each api-request. To hard paste a value there a two easy options:

1: `ctrl`+`c` 🠆 `ctrl`+`alt`+`v` 🠆 under paste, select `values`

2: copy 🠆 right mouse click, under paste select `values`

## Installation
### Get an OpenAI api-key
If you do not have an OpenAI account, create one: [create account](https://auth.openai.com/create-account)

Otherwise, login and go to https://platform.openai.com/api-keys to create a key.

*Note: without a valid `api-key` the function will return an error.*

### Option 1: download macro-enabled example Excel
Easiest option. The Excel is similar to the "Example with cell reference" from above. Modify the Excel to your needs.

### Option 2: download VBA files and install 
1. Download all three files (`OpenAi.bas`, `Json.bas`, and `JsonData.cls`) in the `vba`-folder of the repo
2. Open the Visual Basic Editor (`alt`+`F11`) in Excel
3. In the menu (top left) choose  `File` 🠆 `Import File...`
4. Import <ins>all three</ins> files
5. Follow one of the examples to test if everything works

*Note: you can delete the downloaded files after importing as they are no longer required by Excel.* 



## Credits
This project uses the `vba-json` JSON parser to parse the OpenAI Api response: [mlocati/vba-json](https://github.com/mlocati/vba-json)
