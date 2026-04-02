import { GoogleGenAI } from "@google/genai";

const ai = new GoogleGenAI({ apiKey: process.env.GEMINI_API_KEY });

async function generateGuideImage() {
  const response = await ai.models.generateContent({
    model: 'gemini-2.5-flash-image',
    contents: {
      parts: [
        {
          text: 'A clean, professional infographic showing how to add environment variables in a web development tool. Step 1: Click a gear icon in the top right corner labeled "Settings". Step 2: A sidebar opens with a section called "Environment Variables". Step 3: Two text fields are shown being filled: Key "VITE_SUPABASE_URL" with a URL value, and Key "VITE_SUPABASE_ANON_KEY" with a long alphanumeric string. Use a modern, flat design style with blue and dark grey colors.',
        },
      ],
    },
    config: {
      imageConfig: {
        aspectRatio: "16:9",
        imageSize: "1K"
      }
    }
  });

  for (const part of response.candidates[0].content.parts) {
    if (part.inlineData) {
      return `data:image/png;base64,${part.inlineData.data}`;
    }
  }
}

// Note: This is a conceptual representation of how I would generate the image.
// In the final response, I will provide the text guide and the generated image.
