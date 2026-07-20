import { GoogleGenAI } from '@google/genai';

const ai = new GoogleGenAI({ apiKey: '' });

async function listModels() {
  try {
    const response = await ai.models.list();
    for await (const model of response) {
      console.log(model.name);
    }
  } catch (error) {
    console.error(error);
  }
}

listModels();
