//
//  semaiConfig.js
//  semai
//
//  Created by Yago Arconada on 11/14/25.
//
// SEMAI_OPENAI_API_KEY is defined in secrets.js, which is loaded before this file.
const SEMAI_MODEL = "gpt-4.1-mini"; // or any other chat-capable model id

// Your display name as it appears in Outlook's sender field.
// Used by chat view to identify which messages are yours.
const SEMAI_USER_NAME = "Santiago Arconada Alvarez";

// Style presets for semai
const SEMAI_PRESETS = {
  polite: {
    system:
      "You help me rewrite short email fragments. " +
      "You keep meaning the same, do not add new commitments, and avoid changing facts. " +
      "Make the tone slightly warmer, kind, and professional. " +
      "Keep roughly the same length and never add greetings or signatures.",
    userTemplate:
      "Rewrite the following email fragment in a warm, kind, professional tone. " +
      "Keep it approximately the same length.\n\n\"{{TEXT}}\""
  },
  concise: {
    system:
      "You help me rewrite short email fragments. " +
      "You keep the same meaning and all important details. " +
      "Make the text more concise and direct, but still polite. " +
      "Do not add or remove commitments, and do not add greetings or signatures.",
    userTemplate:
      "Rewrite the following email fragment to be more concise and direct, while staying polite. " +
      "Do not remove key information.\n\n\"{{TEXT}}\""
  },
  custom: {
    system:
      "You help me rewrite short email fragments exactly according to my instruction. " +
      "Preserve the core meaning and facts, do not add new commitments, and avoid changing specific data like dates or numbers.",
    userTemplate:
      "Instruction: {{INSTRUCTION}}\n\n" +
      "Rewrite the following email fragment according to the instruction above:\n\n\"{{TEXT}}\""
  }
};

