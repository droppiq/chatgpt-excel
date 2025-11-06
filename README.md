# ChatGPT for Excel
ChatGPT for Excel using VBA, for those that have an older version of Excel or when Copilot is disabled, e.g. through company policy.

## Usage
It is required to provide a `prompt`, `model`, `reasoning-level`, and `api-key`:
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
>
>Do you want a breakdown by levels too?

*Note: OpenAI can take a few seconds (per query) to return a response. To get a different response, tune the prompt.*

### Example with cell reference
The prompt, model, reasoning, and api-key can also be located in an Excel cell, e.g. cell `A1` holds the prompt, cell `D1` the model, `E1` the reasoning level, and `F1` the OpenAI api-key. To get a response, paste the following in cell `B1`:
```excel
=ChatGPT(A1, $D$1, $E$1, $F$1)
```

*Note the `$` signs in `D1`, `E1`, and `F1` to keep the reference locked if you drag the code down to query the prompts in e.g. cells `A2` to `A10`*

### Model selection (and price per 1M tokens)
Available models and their associated costs can be found on: https://platform.openai.com/docs/pricing

### Reasoning level
It is possible to set the reasoning level for the model to `low`, `medium`, or `high`.

For more information see: https://platform.openai.com/docs/guides/reasoning

## Word of caution
It is **strongly** advised that after a response has been generated to **hard paste** the value.

Option 1: `ctrl` `c` 🠆 `ctrl` `alt` `v` 🠆 under paste, select `values`

Option 2: copy 🠆 right mouse click, under paste select `values`

## Installation

### OpenAI api-key
