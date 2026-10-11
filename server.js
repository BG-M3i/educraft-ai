const express = require("express");
const path = require("path");
const fs = require("fs");
const Database = require("better-sqlite3");
const { Document, Packer, Paragraph, TextRun, HeadingLevel } = require("docx");
const pptxgen = require("pptxgenjs");
const PDFDocument = require("pdfkit");
require("dotenv").config();

const app = express();
const PORT = process.env.PORT || 3000;
const DATA_DIR = process.env.DATA_DIR || path.join(__dirname, "data");
fs.mkdirSync(DATA_DIR, { recursive: true });
const db = new Database(path.join(DATA_DIR, "educraft.sqlite"));
db.exec(`CREATE TABLE IF NOT EXISTS documents (
  id TEXT PRIMARY KEY, title TEXT NOT NULL, type TEXT NOT NULL,
  content TEXT NOT NULL, data TEXT NOT NULL, updated TEXT NOT NULL
)`);
app.use(express.json({ limit: "2mb" }));
app.use(express.static(path.join(__dirname, "public")));

function safeName(s) {
  return String(s || "educraft").replace(/[\\/:*?"<>|]+/g, "-").slice(0, 90) || "educraft";
}
function getDocs() {
  return db.prepare("SELECT * FROM documents ORDER BY updated DESC").all().map(d => ({
    ...d, data: JSON.parse(d.data)
  }));
}
app.get("/api/health", (_req, res) => res.json({
  ok: true, aiConfigured: Boolean(process.env.AI_API_KEY),
  model: process.env.AI_MODEL || "gpt-4o-mini"
}));
app.get("/api/docs", (_req, res) => res.json(getDocs()));
app.post("/api/docs", (req, res) => {
  const { id, title, type, content, data = {} } = req.body || {};
  if (!id || !title || !["lesson", "slides", "test"].includes(type) || typeof content !== "string") {
    return res.status(400).json({ error: "Thiếu hoặc sai thông tin tài liệu." });
  }
  const updated = new Date().toISOString();
  db.prepare(`INSERT INTO documents(id,title,type,content,data,updated) VALUES(?,?,?,?,?,?)
    ON CONFLICT(id) DO UPDATE SET title=excluded.title,type=excluded.type,content=excluded.content,data=excluded.data,updated=excluded.updated`)
    .run(String(id), String(title).slice(0, 200), type, content, JSON.stringify(data), updated);
  res.json({ ok: true, updated });
});
app.delete("/api/docs/:id", (req, res) => {
  db.prepare("DELETE FROM documents WHERE id=?").run(req.params.id);
  res.json({ ok: true });
});

app.post("/api/generate", async (req, res) => {
  const { type, data = {} } = req.body || {};
  if (!["lesson", "slides", "test"].includes(type)) return res.status(400).json({ error: "Loại tài liệu không hợp lệ." });
  if (!process.env.AI_API_KEY) return res.status(503).json({ error: "Chưa cấu hình AI_API_KEY trên máy chủ. Hãy sao chép .env.example thành .env và thêm API key." });
  const labels = { lesson: "kế hoạch bài dạy theo chương trình giáo dục Việt Nam", slides: "dàn ý slide bài giảng 16:9, mỗi slide ngắn gọn kèm ghi chú thuyết trình", test: "đề kiểm tra gồm câu hỏi, thang điểm, đáp án và hướng dẫn chấm tách rõ" };
  const prompt = `Bạn là trợ lý thiết kế học liệu chuyên nghiệp cho giáo viên Việt Nam. Hãy tạo ${labels[type]} bằng tiếng Việt dựa trên thông tin người dùng cung cấp. Trình bày rõ tiêu đề, mục, hoạt động/câu hỏi, thời lượng hoặc điểm số phù hợp. Không bịa nội dung sách cụ thể nếu không được cung cấp; đánh dấu phần cần giáo viên kiểm tra. Nếu thiếu thông tin, dùng giả định hợp lý và ghi chú. Đầu ra chỉ là nội dung văn bản có cấu trúc, không dùng markdown table quá rộng.\n\nThông tin đầu vào:\n${JSON.stringify(data, null, 2)}`;
  try {
    const base = (process.env.AI_BASE_URL || "https://api.openai.com/v1").replace(/\/$/, "");
    const r = await fetch(`${base}/chat/completions`, {
      method: "POST",
      headers: { "Authorization": `Bearer ${process.env.AI_API_KEY}`, "Content-Type": "application/json" },
      body: JSON.stringify({
        model: process.env.AI_MODEL || "gpt-4o-mini",
        messages: [
          { role: "system", content: "Bạn tạo học liệu rõ ràng, chính xác, dễ chỉnh sửa. Không khẳng định nội dung chưa được xác minh." },
          { role: "user", content: prompt }
        ],
        temperature: 0.5
      })
    });
    const json = await r.json();
    if (!r.ok) return res.status(r.status).json({ error: json?.error?.message || "Dịch vụ AI trả về lỗi." });
    const content = json?.choices?.[0]?.message?.content;
    if (!content) return res.status(502).json({ error: "AI không trả về nội dung." });
    res.json({ content });
  } catch (e) {
    res.status(502).json({ error: "Không kết nối được dịch vụ AI. Hãy kiểm tra cấu hình máy chủ." });
  }
});

function linesToParagraphs(content) {
  return String(content || "").split(/\r?\n/).map(line => {
    const t = line.trim();
    if (!t) return new Paragraph({ text: "" });
    if (/^(KẾ HOẠCH|DÀN Ý|ĐỀ KIỂM TRA|SLIDE BÀI GIẢNG)/i.test(t))
      return new Paragraph({ text: t, heading: HeadingLevel.TITLE, spacing: { after: 220 } });
    if (/^(I{1,3}\.|IV\.|PHẦN |HOẠT ĐỘNG |SLIDE \d+|Câu \d+\.?)/i.test(t))
      return new Paragraph({ children: [new TextRun({ text: t, bold: true })], spacing: { before: 100, after: 80 } });
    return new Paragraph({ text: t, spacing: { after: 60 } });
  });
}
app.post("/api/export", async (req, res) => {
  const { format, title, type, content } = req.body || {};
  if (typeof content !== "string" || !["docx", "pptx", "pdf"].includes(format))
    return res.status(400).json({ error: "Định dạng xuất không hợp lệ." });
  const name = safeName(title);
  try {
    if (format === "docx") {
      const doc = new Document({ sections: [{ properties: {}, children: linesToParagraphs(content) }] });
      const buffer = await Packer.toBuffer(doc);
      res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.wordprocessingml.document");
      res.setHeader("Content-Disposition", `attachment; filename*=UTF-8''${encodeURIComponent(name)}.docx`);
      return res.send(buffer);
    }
    if (format === "pptx") {
      const pptx = new pptxgen();
      pptx.layout = "LAYOUT_WIDE";
      pptx.author = "EduCraft AI";
      pptx.subject = "Bài giảng";
      const chunks = content.split(/\n(?=SLIDE \d+:)/).filter(x => /^SLIDE \d+:/.test(x.trim()));
      const source = chunks.length ? chunks : [{ title: name, body: content }];
      for (let i = 0; i < source.length; i++) {
        const block = typeof source[i] === "string" ? source[i] : `${source[i].title}\n${source[i].body}`;
        const lines = block.split(/\r?\n/);
        const slide = pptx.addSlide();
        slide.background = { color: i === 0 ? "EDF2FF" : "FFFFFF" };
        slide.addText((lines.shift() || name).replace(/^SLIDE \d+:\s*/, ""), {
          x: 0.65, y: 0.45, w: 12, h: 0.8, fontFace: "Aptos Display", fontSize: 26,
          bold: true, color: "233B78", margin: 0.05, breakLine: false
        });
        const body = lines.filter(x => x.trim() && !x.startsWith("GHI CHÚ")).join("\n").slice(0, 2500);
        slide.addText(body || block.slice(0, 2500), {
          x: 0.8, y: 1.5, w: 11.5, h: 5.2, fontFace: "Aptos", fontSize: 18,
          color: "243047", breakLine: false, valign: "top", fit: "shrink", margin: 0.12
        });
        slide.addText("EduCraft AI", { x: 0.8, y: 7.08, w: 3, h: 0.2, fontSize: 9, color: "75819A" });
      }
      const buffer = await pptx.write({ outputType: "nodebuffer" });
      res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.presentationml.presentation");
      res.setHeader("Content-Disposition", `attachment; filename*=UTF-8''${encodeURIComponent(name)}.pptx`);
      return res.send(buffer);
    }
    const pdf = new PDFDocument({ margin: 54, size: "A4" });
    res.setHeader("Content-Type", "application/pdf");
    res.setHeader("Content-Disposition", `attachment; filename*=UTF-8''${encodeURIComponent(name)}.pdf`);
    pdf.pipe(res);
    // PDFKit's built-in fonts have limited Vietnamese glyph coverage; use an uploaded font if configured.
    const fontPath = process.env.PDF_FONT_PATH;
    if (fontPath && fs.existsSync(fontPath)) pdf.font(fontPath);
    else pdf.font("Helvetica");
    pdf.fontSize(16).text(name, { align: "center" }).moveDown();
    pdf.fontSize(10.5).text(String(content), { lineGap: 4 });
    pdf.end();
  } catch (e) {
    console.error(e);
    if (!res.headersSent) res.status(500).json({ error: "Xuất file thất bại. Vui lòng thử lại." });
  }
});

app.get("*", (_req, res) => res.sendFile(path.join(__dirname, "public", "index.html")));
app.listen(PORT, () => console.log(`EduCraft AI running on port ${PORT}`));
