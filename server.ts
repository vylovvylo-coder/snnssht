import express from "express";
import { GoogleGenAI } from "@google/genai";
import path from "path";
import { fileURLToPath } from "url";
import { createServer as createViteServer } from "vite";

const __dirname = path.dirname(fileURLToPath(import.meta.url));

const apiKey = process.env.GEMINI_API_KEY || "";
const ai = new GoogleGenAI({
  apiKey: apiKey,
  httpOptions: {
    headers: {
      "User-Agent": "aistudio-build",
    },
  },
});

async function startServer() {
  const app = express();
  app.use(express.json({ limit: "50mb" }));

  // API Routes
  app.post("/api/ai/generate-description", async (req, res) => {
    try {
      const { prompt, category } = req.body;
      const response = await ai.models.generateContent({
        model: "gemini-3.8-flash",
        contents: `أنت خبير تسويق إلكتروني محترف. بناءً على الوصف التالي للمنتج والتصنيف "${category || 'عام'}"، قم بإنشاء:
1. عنوان جذاب للمنتج (بالعربية).
2. وصف تسويقي تفصيلي واحترافي يبرز المزايا والفوائد.
3. 3-5 وسوم (Tags) دلالية.
4. اقتراح سعر تنافسي (رقم تقريبي بالريال السعودي أو الدولار) مع نسبة خصم مقترحة (مثال 10% أو 20%).

أجب حصرياً بصيغة JSON مفاتيحها كالتالي:
{
  "title": "...",
  "description": "...",
  "tags": ["...", "..."],
  "suggestedPrice": 150,
  "discountPercent": 15
}

الوصف الأولي: ${prompt}`,
        config: {
          responseMimeType: "application/json",
        },
      });

      const text = response.text || "{}";
      const data = JSON.parse(text);
      res.json(data);
    } catch (error: any) {
      console.error("AI Description Error:", error);
      res.status(500).json({ error: error.message || "Failed to generate description" });
    }
  });

  app.post("/api/ai/generate-image", async (req, res) => {
    try {
      const { prompt } = req.body;
      const response = await ai.models.generateContent({
        model: "gemini-3.1-flash-lite-image",
        contents: {
          parts: [
            {
              text: `A professional e-commerce product studio photograph, clean minimalist background, ultra realistic, studio lighting: ${prompt}`,
            },
          ],
        },
        config: {
          imageConfig: {
            aspectRatio: "1:1",
          },
        },
      });

      let imageUrl = "";
      if (response.candidates?.[0]?.content?.parts) {
        for (const part of response.candidates[0].content.parts) {
          if (part.inlineData) {
            imageUrl = `data:${part.inlineData.mimeType || 'image/png'};base64,${part.inlineData.data}`;
            break;
          }
        }
      }

      if (!imageUrl) {
        throw new Error("لم يتم توليد أي صورة من النموذج.");
      }

      res.json({ imageUrl });
    } catch (error: any) {
      console.error("AI Image Error:", error);
      res.status(500).json({ error: error.message || "Failed to generate image" });
    }
  });

  // Edit / transform existing image using AI
  app.post("/api/ai/edit-image", async (req, res) => {
    try {
      const { imageBase64, prompt } = req.body;
      const matches = imageBase64.match(/^data:(.+);base64,(.+)$/);
      if (!matches) {
        throw new Error("صيغة الصورة غير صحيحة");
      }
      const mimeType = matches[1];
      const dataBytes = matches[2];

      const response = await ai.models.generateContent({
        model: "gemini-3.1-flash-lite-image",
        contents: {
          parts: [
            {
              inlineData: {
                data: dataBytes,
                mimeType: mimeType,
              },
            },
            {
              text: prompt || "Enhance this product photo with professional studio lighting, clean background, 4k resolution",
            },
          ],
        },
      });

      let imageUrl = "";
      if (response.candidates?.[0]?.content?.parts) {
        for (const part of response.candidates[0].content.parts) {
          if (part.inlineData) {
            imageUrl = `data:${part.inlineData.mimeType || 'image/png'};base64,${part.inlineData.data}`;
            break;
          }
        }
      }

      if (!imageUrl) {
        throw new Error("فشل توليد صورة جديدة من الصورة المدخلة.");
      }

      res.json({ imageUrl });
    } catch (error: any) {
      console.error("AI Edit Image Error:", error);
      res.status(500).json({ error: error.message || "Failed to edit image" });
    }
  });

  app.post("/api/ai/generate-video", async (req, res) => {
    try {
      const { prompt, imageBase64 } = req.body;
      const videoConfig: any = {
        model: "veo-3.1-lite-generate-preview",
        prompt: prompt || "Commercial promotional video for this product with dynamic angles and professional lighting",
        config: {
          numberOfVideos: 1,
          resolution: "720p",
          aspectRatio: "16:9",
        },
      };

      if (imageBase64) {
        const matches = imageBase64.match(/^data:(.+);base64,(.+)$/);
        if (matches) {
          videoConfig.image = {
            mimeType: matches[1],
            imageBytes: matches[2],
          };
        }
      }

      const operation = await ai.models.generateVideos(videoConfig);
      res.json({ operationName: operation.name });
    } catch (error: any) {
      console.error("AI Video Start Error:", error);
      res.status(200).json({
        fallback: true,
        videoUrl: "https://assets.mixkit.co/videos/preview/mixkit-hands-holding-a-smartphone-with-a-green-screen-42938-large.mp4",
        message: "تم تفعيل الفيديو الإعلاني التجريبي (يتطلب توليد الفيديو الفعلي مفتاح API مدفوع)."
      });
    }
  });

  app.post("/api/ai/video-status", async (req, res) => {
    try {
      const { operationName } = req.body;
      const { GenerateVideosOperation } = await import("@google/genai");
      const op = new GenerateVideosOperation();
      op.name = operationName;
      const updated = await ai.operations.getVideosOperation({ operation: op });
      res.json({ done: updated.done, response: updated.response });
    } catch (error: any) {
      console.error("AI Video Status Error:", error);
      res.json({ done: true, fallback: true });
    }
  });

  app.post("/api/ai/video-download", async (req, res) => {
    try {
      const { operationName } = req.body;
      const { GenerateVideosOperation } = await import("@google/genai");
      const op = new GenerateVideosOperation();
      op.name = operationName;
      const updated = await ai.operations.getVideosOperation({ operation: op });
      const uri = updated.response?.generatedVideos?.[0]?.video?.uri;
      if (!uri) {
        return res.status(404).json({ error: "Video URI not found yet" });
      }

      const videoRes = await fetch(uri, {
        headers: { "x-goog-api-key": apiKey },
      });
      
      res.setHeader("Content-Type", "video/mp4");
      const arrayBuffer = await videoRes.arrayBuffer();
      res.send(Buffer.from(arrayBuffer));
    } catch (error: any) {
      console.error("AI Video Download Error:", error);
      res.status(500).json({ error: error.message || "Failed to download video" });
    }
  });

  // Vite middleware for development
  if (process.env.NODE_ENV !== "production") {
    const vite = await createViteServer({
      server: { middlewareMode: true },
      appType: "spa",
    });
    app.use(vite.middlewares);
  } else {
    app.use(express.static(path.join(__dirname, "dist")));
    app.get("*", (req, res) => {
      res.sendFile(path.join(__dirname, "dist", "index.html"));
    });
  }

  const port = Number(process.env.PORT || 3000);
  app.listen(port, "0.0.0.0", () => {
    console.log(`Server running on http://localhost:${port}`);
  });
}

startServer();
