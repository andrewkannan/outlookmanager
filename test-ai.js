import { GoogleGenAI } from '@google/genai';

async function test() {
    try {
        const ai = new GoogleGenAI({ apiKey: '' });
        const response = await ai.models.generateContent({
            model: 'gemini-2.5-flash',
            contents: 'Hello, this is a test.',
        });
        console.log("SUCCESS:", response.text);
    } catch (e) {
        console.error("ERROR:", e.message);
    }
}
test();
