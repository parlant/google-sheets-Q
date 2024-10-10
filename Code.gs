/**
 * Q – AI Function for Google Sheets™
 * https://parlant.xyz/l/google-sheets-chatgpt-function
 * 
 * MIT LICENSE
 * 
 * Copyright 2024 Parlant GmbH
 * 
 * Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the “Software”), to deal in the Software without restriction, including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED “AS IS”, WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
 */

const OPENAI_PREFERRED_MODELS = [
  'gpt-4o',
  'gpt-4o-mini',
  'gpt-4-turbo',
  'gpt-3.5-turbo',
];
const OPENAI_REASONING_MODEL_REGEX = /^o1-(mini|preview).*/;
const OPENAI_DEFAULT_SYSTEM = 'You are a helpful assistant.';
const OPENAI_DEFAULT_TEMPERATURE = 0.7;
const OPENAI_MAX_TOKENS = 2048;
const CACHE_TIMEOUT = 21600;

function onOpen(event) {
  createMenu();
}

function createMenu() {
  SpreadsheetApp
    .getUi()
    .createAddonMenu()
    .addItem('Set OpenAI API key', 'showApiKeyPrompt')
    .addToUi();
}

function showApiKeyPrompt(prompt) {
  if (prompt === undefined) {
    prompt = 'OpenAI API key';
  }
  const ui = SpreadsheetApp.getUi();
  const response = ui.prompt(prompt, 'Enter your OpenAI API key here', ui.ButtonSet.OK_CANCEL);
  if (response.getSelectedButton() == ui.Button.OK) {
    const documentProperties = PropertiesService.getDocumentProperties();
    documentProperties.setProperty('OPENAI_API_KEY', response.getResponseText());
    setDefaultModel(true);
  }
}

function getDefaultModel() {
  const documentProperties = PropertiesService.getDocumentProperties();
  return documentProperties.getProperty('OPENAI_DEFAULT_MODEL');
}

function setDefaultModel( ) {
  const documentProperties = PropertiesService.getDocumentProperties();
  const api_key = documentProperties.getProperty('OPENAI_API_KEY');
  if (!api_key) {
    throw new Error('OPENAI_API_KEY is not set, please set it using the menu "Extensions" > "Q" > "Set OpenAI API key"');
  }
  const url = 'https://api.openai.com/v1/models';
  const response = UrlFetchApp.fetch(url, {
    method: 'get',
    contentType: 'application/json',
    headers: { Authorization: `Bearer ${api_key}` }
  });
  const result = JSON.parse(response.getContentText());
  const availableModels = result.data.map(e => e.id);
  const preferredDefaultModel = OPENAI_PREFERRED_MODELS.find(preferredModel => availableModels.find(model => preferredModel == model));
  documentProperties.setProperty('OPENAI_DEFAULT_MODEL', preferredDefaultModel);
  return preferredDefaultModel;
}

/**
 * This function lets you query OpenAI.
 * 
 * @param {string} prompt - The user message sent to OpenAI.
 * @param {string} [system_prompt] - [OPTIONAL] The system message sent to OpenAI, defaults to 'You are a helpful assistant.'.
 * @param {string} [model] - [OPTIONAL] The model to answer prompts with, defaults to gpt-4o.
 * @param {number} [temperature] - [OPTIONAL] The temperature param, defaults to 0.7.
 * @param {number} [max_tokens] - [OPTIONAL] The maximum tokens param, defaults to 2048.
 * @param {number|boolean} [cache_timeout] - [OPTIONAL] The cache timeout that responses for the exact same requests are stored, defaults to 21600 (6 hours). Set to 0 or false to disable caching.
 * @return The text response from OpenAI.
 * @customfunction
 */
function Q(prompt, system_prompt, model, temperature, max_tokens, cache_timeout) {
  const api_key = PropertiesService.getDocumentProperties().getProperty('OPENAI_API_KEY');
  if (!api_key) {
    throw new Error('OPENAI_API_KEY is not set, please set it using the menu "Extensions" > "Q" > "Set OpenAI API key"');
  }
  if (model === undefined || model === null) {
    model = getDefaultModel();
    if (!model) {
      model = setDefaultModel();
    }
  }
  if (system_prompt === undefined || system_prompt === null) {
    system_prompt = OPENAI_DEFAULT_SYSTEM;
  }
  let is_reasoning_model = !!model.match(OPENAI_REASONING_MODEL_REGEX);
  if (temperature === undefined || temperature === null) {
    temperature = OPENAI_DEFAULT_TEMPERATURE;
  }
  if (max_tokens === undefined || max_tokens === null) {
    max_tokens = OPENAI_MAX_TOKENS;
  }
  if (cache_timeout === undefined) {
    cache_timeout = CACHE_TIMEOUT;
  }
  let cache_buster = false;
  if (typeof cache_timeout === 'number' && cache_timeout > CACHE_TIMEOUT) {
    cache_timeout = CACHE_TIMEOUT;
  }
  if (typeof cache_timeout === 'number' && cache_timeout < 0) {
    cache_timeout = 0;
  }
  if (cache_timeout === 0 || cache_timeout === false) {
    cache_buster = Utilities.getUuid();
  }
  const cache = CacheService.getDocumentCache();
  const cacheKeyResult = Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, `${prompt}${system_prompt}${model}${temperature}${max_tokens}${cache_buster}`);
  const cacheKey = Utilities.base64EncodeWebSafe(cacheKeyResult);
  const cachedAnswer = cache.get(cacheKey);
  if (cachedAnswer) {
    console.info(`Using response from cache: ${cacheKey}`);
    return cachedAnswer;
  }
  const url = 'https://api.openai.com/v1/chat/completions';
  let messages = [];
  if (!is_reasoning_model) {
    messages.push({"role": "system", "content": system_prompt});
  }
  messages.push({"role": "user", "content": prompt});
  const data = {
    model: model,
    messages: messages,
  };
  if (is_reasoning_model) {
    data['max_completion_tokens'] = max_tokens;
  } else {
    data['max_tokens'] = max_tokens;
    data['temperature'] = temperature;
  }
  const response = UrlFetchApp.fetch(url, {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(data),
    headers: { Authorization: `Bearer ${api_key}` }
  });
  const result = JSON.parse(response.getContentText());
  let answer = '';
  for (i in result.choices) {
    answer += result.choices[i].message.content;
  }
  if (!cache_buster) {
    console.info(`Storing response in cache: ${cacheKey}`);
    cache.put(cacheKey, answer, cache_timeout);
  }
  return answer;
}

function test_Q() {
  const answer = Q(`How are you today? Your UUID for today is: ${Utilities.getUuid()}`, "You are a helpful assistant, always include your UUID if I send you one.");
  console.log(answer);
  return answer;
}

function test_setDefaultModel() {
  setDefaultModel();
  const documentProperties = PropertiesService.getDocumentProperties();
  const defaultModel = documentProperties.getProperty('OPENAI_DEFAULT_MODEL');
  console.log(`OPENAI_DEFAULT_MODEL: ${defaultModel}`);
}

function test_reasoningModel() {
  const answer = Q(`How are you today? Your UUID for today is: ${Utilities.getUuid()}`, "You are a helpful assistant, always include your UUID if I send you one.", "o1-mini");
  console.log(answer);
  return answer;
}
