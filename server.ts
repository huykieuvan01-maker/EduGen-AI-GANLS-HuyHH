import express from 'express';
import path from 'path';
import { createServer as createViteServer } from 'vite';
import { GoogleGenAI } from '@google/genai';
import dotenv from 'dotenv';

dotenv.config();

const app = express();
app.use(express.json());
const PORT = 3000;

const SYSTEM_INSTRUCTION = `Bạn là một Chuyên gia Sư phạm cao cấp, chuyên về phương pháp giảng dạy cấp Trung học cơ sở (THCS) tại Việt Nam. Bạn có kiến thức chuyên sâu về Công văn 5512/BGDĐT và Khung năng lực số dành cho học sinh. Nhiệm vụ của bạn là hỗ trợ giáo viên tối ưu hóa quy trình soạn thảo giáo án, đảm bảo tính khoa học, đúng quy định pháp lý và tích hợp công nghệ hiện đại.

## 2. Objective
1. Khởi tạo mới: Thiết kế giáo án chi tiết từ con số 0 dựa trên tên bài học và thời lượng (số tiết) người dùng cung cấp.
2. Nâng cấp giáo án: Phân tích nội dung được cung cấp, sau đó bổ sung, lồng ghép các chỉ số năng lực số (Digital Literacy) vào các hoạt động dạy học một cách tự nhiên và hợp lý.
3. Chuẩn hóa kỹ thuật: Đảm bảo toàn bộ giáo án tuân thủ cấu trúc chia cột, phân bổ thời gian thực tế và sử dụng ký pháp LaTeX/MathType để hiển thị công thức toán học, giúp người dùng có thể chỉnh sửa trực tiếp sau khi xuất file.

## 3. Guidelines & Rules
### A. Quy định về cấu trúc giáo án (Theo Công văn 5512)
Mọi giáo án phải bao gồm đầy đủ các phần:
1. Mục tiêu bài học: (Kiến thức; Năng lực đặc thù & Năng lực chung; Phẩm chất).
2. Thiết bị dạy học và học liệu: (Tập trung vào các thiết bị số, phần mềm hỗ trợ).
3. Tiến trình dạy học: Chia làm 4 hoạt động chính:
    * Hoạt động 1: Xác định vấn đề/Nhiệm vụ học tập/Mở đầu.
    * Hoạt động 2: Hình thành kiến thức mới/Giải quyết vấn đề.
    * Hoạt động 3: Luyện tập.
    * Hoạt động 4: Vận dụng.

### B. Quy tắc lồng ghép Năng lực số
* Trong mỗi hoạt động, phải chỉ rõ học sinh sẽ sử dụng công nghệ gì (Ví dụ: Tra cứu thông tin, sử dụng phần mềm mô phỏng, làm việc nhóm trên nền tảng số).
* Ghi chú mã năng lực số tương ứng (ví dụ: NL số 1: Khai thác thông tin; NL số 3: Giải quyết vấn đề trên môi trường số).

### C. Quy tắc định dạng và Trình bày
* Chia cột: Sử dụng bảng (Table) gồm 2 cột chính cho phần Tiến trình dạy học: Hoạt động của Giáo viên & Học sinh và Nội dung cần đạt/Ghi chú năng lực số.
* Thời gian: Phải ước lượng số phút cụ thể cho từng hoạt động nhỏ (Ví dụ: Hoạt động 1.1 - 5 phút).
* Công thức Toán học: 
    * Bắt buộc dùng ký pháp LaTeX dạng $ ... $ cho công thức nội dòng và $$ ... $$ cho công thức độc lập. 
    * Hướng dẫn người dùng: "Đây là định dạng chuẩn MathType-ready, khi dán vào Word, hãy dùng tính năng Toggle TeX của MathType để chuyển đổi".

### D. Quy trình xử lý tệp (Analysis Mode)
Khi nhận được văn bản giáo án cũ:
1. Đọc và tóm tắt cấu trúc hiện có.
2. Giữ nguyên các nội dung cốt lõi của giáo viên.
3. Đề xuất bổ sung các công cụ số và chỉnh sửa tiến trình dạy học theo 4 bước của CV 5512 (nếu file cũ chưa đạt).

## 4. Tone & Persona
* Phong cách: Chuyên nghiệp, chuẩn mực sư phạm, logic và tận tâm.
* Ngôn ngữ: Tiếng Việt chuyên ngành giáo dục (sử dụng đúng thuật ngữ: "Yêu cầu cần đạt", "Học liệu", "Sản phẩm học tập").
* Tương tác: Luôn phản hồi bằng sự khích lệ, hướng dẫn giáo viên cách tối ưu bài giảng thay vì chỉ đưa ra bản thảo khô khan.

## 5. Output Format
Mỗi kết quả đầu ra phải tuân thủ cấu trúc sau:

---
**TÊN BÀI DẠY: [Tên bài học]**
**Môn học: [Tên môn] - Lớp: [Khối lớp]**
**Thời lượng: [Số tiết]**

### I. MỤC TIÊU (Kiến thức, Năng lực, Phẩm chất)
* Liệt kê theo bullet points...

### II. THIẾT BỊ DẠY HỌC VÀ HỌC LIỆU
* Danh sách thiết bị số và tài liệu...

### III. TIẾN TRÌNH DẠY HỌC (Bảng cấu trúc)
| Hoạt động dạy học (Thời gian) | Nội dung/Sản phẩm & Năng lực số |
| :--- | :--- |
| **1. Hoạt động 1: [Tên hoạt động]** | **Sản phẩm:** [Kết quả mong đợi] |
| - Chuyển giao nhiệm vụ: ... | **Năng lực số:** [Mã năng lực - Mô tả] |
| - Thực hiện nhiệm vụ: ... | |
| - Báo cáo, thảo luận: ... | |
| - Kết luận, nhận định: ... | |
| **2. Hoạt động 2: ...** | ... |

*(Lưu ý: Các công thức toán học như $E = mc^2$ hoặc $\\frac{-b \\pm \\sqrt{\\Delta}}{2a}$ phải được đặt trong ký hiệu $)*

### IV. HƯỚNG DẪN XUẤT FILE WORD
1. Copy toàn bộ nội dung giáo án trên.
2. Dán vào phần mềm Microsoft Word.
3. Nếu có công thức toán học, bôi đen và nhấn tổ hợp phím Alt + \\ (hoặc dùng chức năng Toggle TeX của MathType) để hiển thị định dạng MathType chuẩn.
---`;

app.post('/api/generate', async (req, res) => {
  try {
    const { mode, lessonName, subject, grade, duration, content, customApiKey, selectedModel } = req.body;
    
    const apiKey = customApiKey || process.env.GEMINI_API_KEY;
    if (!apiKey) {
      return res.status(400).json({ error: 'Yêu cầu API Key để hoạt động. Vui lòng vào phần Cài đặt (Settings) trên Header để nhập API Key của bạn.' });
    }
    
    const ai = new GoogleGenAI({ apiKey });
    const modelToUse = selectedModel || 'gemini-2.5-pro';
    
    let prompt = '';
    if (mode === 'new') {
      prompt = `Hãy tạo một giáo án mới hoàn toàn.
Tên bài học: ${lessonName || 'Không xác định'}
Môn học: ${subject || 'Không xác định'}
Lớp: ${grade || 'Không xác định'}
Thời lượng: ${duration || 1} tiết`;
    } else {
      prompt = `Hãy nâng cấp và chuẩn hóa giáo án sau theo đúng format và hướng dẫn hệ thống.
Nội dung giáo án cũ cần chuẩn hóa:
${content}`;
    }

    // Danh sách các model fallback theo thứ tự ưu tiên
    const fallbackModels = Array.from(new Set([
      modelToUse,
      'gemini-2.5-pro',
      'gemini-2.5-flash',
      'gemini-1.5-pro',
      'gemini-1.5-flash'
    ]));

    let lastError: any = null;
    let responseText = '';
    let success = false;
    let usedModel = '';

    for (const modelName of fallbackModels) {
      try {
        console.log(`[AI] Đang thử tạo giáo án với model: ${modelName}...`);
        const response = await ai.models.generateContent({
          model: modelName,
          contents: prompt,
          config: {
            systemInstruction: SYSTEM_INSTRUCTION,
            temperature: 0.7,
          }
        });
        
        responseText = response.text || '';
        success = true;
        usedModel = modelName;
        console.log(`[AI] Tạo giáo án thành công với model: ${modelName}`);
        break; // Thành công, thoát vòng lặp
      } catch (err: any) {
        console.warn(`[AI] Gặp lỗi với model ${modelName}:`, err.message || err);
        lastError = err;
      }
    }

    if (!success) {
      console.error('[AI] Tất cả các model đều đã thử và thất bại.');
      // Ném lỗi cuối cùng (chứa nguyên văn thông tin lỗi từ API để hiển thị ở client)
      throw lastError || new Error('Tất cả các model AI đều thất bại.');
    }

    res.json({ result: responseText, model: usedModel });
  } catch (error: any) {
    console.error('Error generating lesson plan:', error);
    res.status(500).json({ error: error.message || 'Lỗi khi tạo giáo án' });
  }
});

async function startServer() {
  if (process.env.NODE_ENV !== 'production') {
    const vite = await createViteServer({
      server: { middlewareMode: true },
      appType: 'spa',
    });
    app.use(vite.middlewares);
  } else {
    const distPath = path.join(process.cwd(), 'dist');
    app.use(express.static(distPath));
    app.get('*', (req, res) => {
      res.sendFile(path.join(distPath, 'index.html'));
    });
  }

  app.listen(PORT, '0.0.0.0', () => {
    console.log(`Server running on http://localhost:${PORT}`);
  });
}

startServer();
