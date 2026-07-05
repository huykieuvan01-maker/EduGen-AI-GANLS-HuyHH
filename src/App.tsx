import React, { useState } from 'react';
import ReactMarkdown from 'react-markdown';
import remarkGfm from 'remark-gfm';
import remarkMath from 'remark-math';
import rehypeKatex from 'rehype-katex';
import { BookOpen, FileText, Sparkles, Copy, Check, Loader2, ArrowRight, UploadCloud, X, Settings, Download } from 'lucide-react';
import 'katex/dist/katex.min.css';
import { cn } from './lib/utils';

// Khai báo các thư viện được nhúng từ CDN
declare const mammoth: any;
declare const pdfjsLib: any;

export default function App() {
  const [mode, setMode] = useState<'new' | 'upgrade'>('new');
  const [lessonName, setLessonName] = useState('');
  const [subject, setSubject] = useState('');
  const [grade, setGrade] = useState('');
  const [duration, setDuration] = useState('');
  const [content, setContent] = useState('');
  
  const [result, setResult] = useState<string | null>(null);
  const [isLoading, setIsLoading] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const [copied, setCopied] = useState(false);

  const [fileLoading, setFileLoading] = useState(false);
  const [fileError, setFileError] = useState<string | null>(null);
  const [fileName, setFileName] = useState<string | null>(null);
  const [isDragging, setIsDragging] = useState(false);
  const [customApiKey, setCustomApiKey] = useState(() => localStorage.getItem('gemini_api_key') || '');
  const [selectedModel, setSelectedModel] = useState(() => localStorage.getItem('gemini_model') || 'gemini-2.5-flash');
  const [showSettings, setShowSettings] = useState(() => !localStorage.getItem('gemini_api_key'));

  const saveSettings = (key: string, model: string) => {
    localStorage.setItem('gemini_api_key', key);
    localStorage.setItem('gemini_model', model);
    setCustomApiKey(key);
    setSelectedModel(model);
    setShowSettings(false);
  };

  const exportToDocx = (markdown: string, titleName: string = "Giao_an") => {
    if (typeof (window as any).docx === 'undefined') {
      alert("Thư viện docx.js chưa được tải hoàn tất từ CDN. Vui lòng kiểm tra lại kết nối mạng của bạn.");
      return;
    }

    const { Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, 
            HeadingLevel, AlignmentType, WidthType, BorderStyle } = (window as any).docx;

    const lines = markdown.split('\n');
    const docChildren: any[] = [];
    
    // Borders style cho Cell
    const cellBorders = {
      top: { style: BorderStyle.SINGLE, size: 4, color: "CBD5E1" },
      bottom: { style: BorderStyle.SINGLE, size: 4, color: "CBD5E1" },
      left: { style: BorderStyle.SINGLE, size: 4, color: "CBD5E1" },
      right: { style: BorderStyle.SINGLE, size: 4, color: "CBD5E1" },
    };

    let inTable = false;
    let tableHeaders: string[] = [];
    let tableRowsData: string[][] = [];

    const flushTable = () => {
      if (tableRowsData.length === 0) return;
      
      const docRows = [];
      
      // Header Row
      if (tableHeaders.length > 0) {
        docRows.push(
          new TableRow({
            children: tableHeaders.map(headerText => 
              new TableCell({
                children: [
                  new Paragraph({
                    children: [new TextRun({ text: headerText.trim(), bold: true, size: 24, font: "Times New Roman" })],
                    alignment: AlignmentType.CENTER,
                    spacing: { before: 120, after: 120 }
                  })
                ],
                shading: { fill: "F1F5F9" },
                borders: cellBorders,
                width: { size: 50, type: WidthType.PERCENTAGE },
                verticalAlign: (window as any).docx.VerticalAlign.CENTER
              })
            ),
            tableHeader: true
          })
        );
      }
      
      // Body Rows
      tableRowsData.forEach(rowData => {
        // Kiểm tra xem hàng này có phải là hàng phân mục chính hoặc hoạt động lớn (colspan 2)
        const cleanCol1 = rowData[0]?.trim() || '';
        const cleanCol2 = rowData[1]?.trim() || '';
        const isActivityTitle = cleanCol1.startsWith('**') && (cleanCol2 === '' || cleanCol2 === '-');
        
        docRows.push(
          new TableRow({
            children: rowData.map((cellText, cellIndex) => {
              const cleanText = cellText.trim();
              
              // Nếu hàng là tiêu đề hoạt động, chúng ta gộp cột (colspan = 2) cho cell đầu tiên và bỏ qua cell thứ 2
              if (isActivityTitle && cellIndex > 0) return null as any;

              // Parse inline text với các dòng ngắt bởi <br> hoặc \n
              const paragraphsInCell = cleanText.split('<br>').map(paraText => {
                const runs: any[] = [];
                const regex = /\*\*(.*?)\*\*/g;
                let match;
                let lastIndex = 0;
                
                while ((match = regex.exec(paraText)) !== null) {
                  const textBefore = paraText.substring(lastIndex, match.index);
                  if (textBefore) {
                    runs.push(new TextRun({ text: textBefore.replace(/\\n/g, '\n'), size: 24, font: "Times New Roman" }));
                  }
                  runs.push(new TextRun({ text: match[1], bold: true, size: 24, font: "Times New Roman" }));
                  lastIndex = regex.lastIndex;
                }
                
                const textRemaining = paraText.substring(lastIndex);
                if (textRemaining) {
                  runs.push(new TextRun({ text: textRemaining.replace(/\\n/g, '\n'), size: 24, font: "Times New Roman" }));
                }
                
                if (runs.length === 0 && paraText) {
                  runs.push(new TextRun({ text: paraText.replace(/\\n/g, '\n'), size: 24, font: "Times New Roman" }));
                }

                return new Paragraph({
                  children: runs,
                  spacing: { before: 80, after: 80 },
                  alignment: isActivityTitle ? AlignmentType.CENTER : AlignmentType.LEFT
                });
              });

              return new TableCell({
                children: paragraphsInCell.length > 0 ? paragraphsInCell : [new Paragraph({ children: [] })],
                borders: cellBorders,
                columnSpan: isActivityTitle ? 2 : undefined,
                shading: isActivityTitle ? { fill: "F8FAFC" } : undefined,
                width: isActivityTitle ? { size: 100, type: WidthType.PERCENTAGE } : { size: cellIndex === 0 ? 45 : 55, type: WidthType.PERCENTAGE },
                verticalAlign: (window as any).docx.VerticalAlign.TOP
              });
            }).filter(cell => cell !== null)
          })
        );
      });
      
      const table = new Table({
        rows: docRows,
        width: { size: 100, type: WidthType.PERCENTAGE }
      });
      
      docChildren.push(table);
      docChildren.push(new Paragraph({ spacing: { before: 200 } })); // Khoảng cách sau bảng
      
      tableRowsData = [];
      tableHeaders = [];
      inTable = false;
    };

    for (let i = 0; i < lines.length; i++) {
      const line = lines[i].trim();
      
      // Phát hiện bảng Markdown
      if (line.startsWith('|')) {
        if (line.includes(':---') || line.includes('---:')) {
          inTable = true;
          continue;
        }
        
        const parts = line.split('|').map(p => p.trim());
        if (parts[0] === '') parts.shift();
        if (parts[parts.length - 1] === '') parts.pop();
        
        if (!inTable) {
          tableHeaders = parts;
          inTable = true;
        } else {
          if (parts.length === 1) {
            parts.push(''); // Thêm cột trống để tránh lỗi colspan
          }
          tableRowsData.push(parts);
        }
        continue;
      } else {
        if (inTable) {
          flushTable();
        }
      }

      if (line === '') {
        docChildren.push(new Paragraph({ spacing: { before: 120 } }));
        continue;
      }

      // Tiêu đề giáo án lớn (ví dụ: **TÊN BÀI DẠY...** hoặc # Tên Bài)
      if (line.startsWith('# ') || (line.startsWith('**') && line.endsWith('**') && (line.includes('BÀI DẠY') || line.includes('MÔN') || line.includes('THỜI LƯỢNG')))) {
        const text = line.replace(/#/g, '').replace(/\*\*/g, '').trim();
        docChildren.push(
          new Paragraph({
            children: [new TextRun({ text, bold: true, size: 28, font: "Times New Roman" })],
            alignment: AlignmentType.CENTER,
            spacing: { before: 240, after: 120 }
          })
        );
        continue;
      }

      // Các tiêu đề chương mục lớn (### I., ### II., ### III., ### IV.)
      if (line.startsWith('### ')) {
        const text = line.substring(4).replace(/\*\*/g, '').trim();
        docChildren.push(
          new Paragraph({
            children: [new TextRun({ text, bold: true, size: 26, font: "Times New Roman" })],
            heading: HeadingLevel.HEADING_2,
            spacing: { before: 180, after: 120 }
          })
        );
        continue;
      }

      // Các đề mục nhỏ hơn (ví dụ * 1. Hoạt động hoặc **1. Hoạt động**)
      if (line.startsWith('* **') || line.startsWith('**')) {
        const text = line.replace(/^\*\s+/, '').replace(/\*\*/g, '').trim();
        docChildren.push(
          new Paragraph({
            children: [new TextRun({ text, bold: true, size: 24, font: "Times New Roman" })],
            spacing: { before: 120, after: 60 }
          })
        );
        continue;
      }

      // Các dòng bullet points thông thường (bắt đầu bằng * hoặc -)
      if (line.startsWith('*') || line.startsWith('-')) {
        const text = line.substring(1).trim();
        const runs = [];
        let lastIdx = 0;
        const regex = /\*\*(.*?)\*\*/g;
        let match;
        
        while ((match = regex.exec(text)) !== null) {
          const textBefore = text.substring(lastIdx, match.index);
          if (textBefore) {
            runs.push(new TextRun({ text: textBefore, size: 24, font: "Times New Roman" }));
          }
          runs.push(new TextRun({ text: match[1], bold: true, size: 24, font: "Times New Roman" }));
          lastIdx = regex.lastIndex;
        }
        
        const textRemaining = text.substring(lastIdx);
        if (textRemaining) {
          runs.push(new TextRun({ text: textRemaining, size: 24, font: "Times New Roman" }));
        }
        
        if (runs.length === 0) {
          runs.push(new TextRun({ text, size: 24, font: "Times New Roman" }));
        }

        docChildren.push(
          new Paragraph({
            children: runs,
            bullet: { level: 0 },
            spacing: { before: 60, after: 60 }
          })
        );
        continue;
      }

      // Dòng văn bản bình thường
      const runs = [];
      let lastIdx = 0;
      const regex = /\*\*(.*?)\*\*/g;
      let match;
      
      while ((match = regex.exec(line)) !== null) {
        const textBefore = line.substring(lastIdx, match.index);
        if (textBefore) {
          runs.push(new TextRun({ text: textBefore, size: 24, font: "Times New Roman" }));
        }
        runs.push(new TextRun({ text: match[1], bold: true, size: 24, font: "Times New Roman" }));
        lastIdx = regex.lastIndex;
      }
      
      const textRemaining = line.substring(lastIdx);
      if (textRemaining) {
        runs.push(new TextRun({ text: textRemaining, size: 24, font: "Times New Roman" }));
      }
      
      if (runs.length === 0) {
        runs.push(new TextRun({ text: line, size: 24, font: "Times New Roman" }));
      }

      docChildren.push(
        new Paragraph({
          children: runs,
          spacing: { before: 60, after: 60 }
        })
      );
    }

    if (inTable) {
      flushTable();
    }

    const doc = new Document({
      styles: {
        default: {
          document: {
            run: { font: "Times New Roman", size: 24 }
          }
        }
      },
      sections: [{
        properties: {
          page: {
            margin: {
              top: 1134,    // 2cm
              bottom: 1134, // 2cm
              left: 1701,   // 3cm
              right: 850    // 1.5cm
            }
          }
        },
        children: docChildren
      }]
    });

    Packer.toBlob(doc).then((blob: any) => {
      const url = window.URL.createObjectURL(blob);
      const a = document.createElement("a");
      a.href = url;
      a.download = `${titleName.replace(/[\s/\\:*?"<>|]/g, "_") || "Giao_an_5512"}.docx`;
      document.body.appendChild(a);
      a.click();
      document.body.removeChild(a);
      window.URL.revokeObjectURL(url);
    }).catch((err: any) => {
      console.error("Lỗi khi tạo file Word:", err);
      alert("Đã xảy ra lỗi khi xuất file Word: " + err.message);
    });
  };

  const parsePdf = async (file: File): Promise<string> => {
    return new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.onload = async (e) => {
        try {
          const typedarray = new Uint8Array(e.target?.result as ArrayBuffer);
          
          if (typeof pdfjsLib === 'undefined') {
            throw new Error('Thư viện PDF.js chưa được tải. Vui lòng kiểm tra kết nối mạng của bạn.');
          }

          // Cấu hình đường dẫn worker từ CDN để xử lý parsing
          pdfjsLib.GlobalWorkerOptions.workerSrc = 'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js';

          const pdf = await pdfjsLib.getDocument({ data: typedarray }).promise;
          let fullText = '';
          
          for (let i = 1; i <= pdf.numPages; i++) {
            const page = await pdf.getPage(i);
            const textContent = await page.getTextContent();
            const pageText = textContent.items.map((item: any) => item.str).join(' ');
            fullText += pageText + '\n';
          }
          
          resolve(fullText.trim());
        } catch (err: any) {
          reject(new Error(err.message || 'Lỗi khi giải nén file PDF'));
        }
      };
      reader.onerror = () => reject(new Error('Lỗi đọc file PDF'));
      reader.readAsArrayBuffer(file);
    });
  };

  const parseDocx = async (file: File): Promise<string> => {
    return new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.onload = async (e) => {
        try {
          const arrayBuffer = e.target?.result as ArrayBuffer;
          if (typeof mammoth === 'undefined') {
            throw new Error('Thư viện Mammoth chưa được tải. Vui lòng kiểm tra kết nối mạng của bạn.');
          }
          const result = await mammoth.extractRawText({ arrayBuffer });
          resolve(result.value.trim());
        } catch (err: any) {
          reject(new Error(err.message || 'Lỗi khi giải nén file Word'));
        }
      };
      reader.onerror = () => reject(new Error('Lỗi đọc file Word'));
      reader.readAsArrayBuffer(file);
    });
  };

  const handleFileChange = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;
    await processFile(file);
  };

  const processFile = async (file: File) => {
    setFileLoading(true);
    setFileError(null);
    setFileName(file.name);

    try {
      const extension = file.name.split('.').pop()?.toLowerCase();
      let extractedText = '';

      if (extension === 'docx') {
        extractedText = await parseDocx(file);
      } else if (extension === 'pdf') {
        extractedText = await parsePdf(file);
      } else {
        throw new Error('Định dạng file không hỗ trợ. Vui lòng tải file .docx hoặc .pdf');
      }

      if (!extractedText) {
        throw new Error('Không thể đọc nội dung văn bản từ tệp này hoặc tệp trống.');
      }

      setContent(extractedText);
    } catch (err: any) {
      console.error(err);
      setFileError(err.message || 'Lỗi khi xử lý file');
      setFileName(null);
    } finally {
      setFileLoading(false);
    }
  };

  const handleDragOver = (e: React.DragEvent) => {
    e.preventDefault();
    setIsDragging(true);
  };

  const handleDragLeave = () => {
    setIsDragging(false);
  };

  const handleDrop = async (e: React.DragEvent) => {
    e.preventDefault();
    setIsDragging(false);
    
    const file = e.dataTransfer.files?.[0];
    if (file) {
      await processFile(file);
    }
  };

  const handleGenerate = async () => {
    setIsLoading(true);
    setError(null);
    setResult(null);

    try {
      let responseText = '';
      let usedModel = selectedModel;

      // Nếu người dùng đã tự nhập API Key cá nhân, gọi trực tiếp từ client-side
      // Điều này giúp tránh hoàn toàn giới hạn Vercel 10s Timeout đối với các giáo án dài!
      if (customApiKey && customApiKey.trim()) {
        console.log("[Client] Calling Gemini API directly...");
        const geminiUrl = `https://generativelanguage.googleapis.com/v1beta/models/${selectedModel}:generateContent?key=${customApiKey.trim()}`;
        
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

        const systemInstructionText = `Bạn là một Chuyên gia Sư phạm cao cấp, chuyên về phương pháp giảng dạy cấp Trung học cơ sở (THCS) tại Việt Nam. Bạn có kiến thức chuyên sâu về Công văn 5512/BGDĐT và Khung năng lực số dành cho học sinh. Nhiệm vụ của bạn là hỗ trợ giáo viên tối ưu hóa quy trình soạn thảo giáo án, đảm bảo tính khoa học, đúng quy định pháp lý và tích hợp công nghệ hiện đại.

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

*(Lưu ý: Các công thức toán học như $E = mc^2$ hoặc \\frac{-b \\pm \\sqrt{\\Delta}}{2a} phải được đặt trong ký hiệu $)*

### IV. HƯỚNG DẪN XUẤT FILE WORD
1. Copy toàn bộ nội dung giáo án trên.
2. Dán vào phần mềm Microsoft Word.
3. Nếu có công thức toán học, bôi đen và nhấn tổ hợp phím Alt + \\ (hoặc dùng chức năng Toggle TeX của MathType) để hiển thị định dạng MathType chuẩn.
---`;

        const res = await fetch(geminiUrl, {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({
            contents: [{ parts: [{ text: prompt }] }],
            systemInstruction: { parts: [{ text: systemInstructionText }] },
            generationConfig: { temperature: 0.7 }
          })
        });

        const resData = await res.json();
        
        if (!res.ok) {
          throw new Error(resData.error?.message || `Lỗi từ Gemini API (Mã: ${res.status})`);
        }

        responseText = resData.candidates?.[0]?.content?.parts?.[0]?.text || '';
        if (!responseText) {
          throw new Error("Phản hồi trống từ Gemini API.");
        }
      } else {
        // Fallback: Nếu không có key cá nhân, gọi qua Vercel serverless function (dùng server key)
        console.log("[Client] Calling backend function...");
        const res = await fetch('/api/generate', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({ 
            mode, 
            lessonName, 
            subject, 
            grade, 
            duration, 
            content,
            selectedModel
          }),
        });

        let data: any = {};
        const contentType = res.headers.get("content-type");
        if (contentType && contentType.includes("application/json")) {
          data = await res.json();
        } else {
          const textError = await res.text();
          console.error("Non-JSON response:", textError);
          throw new Error(`Phản hồi không hợp lệ từ máy chủ (Mã: ${res.status}). Có thể giáo án đầu vào quá dài khiến server bị quá tải thời gian xử lý (Vercel Timeout). Bạn nên tự cài đặt API Key cá nhân trong "Cấu hình AI" ở Header để chạy trực tiếp không bị giới hạn timeout.`);
        }
        
        if (!res.ok) {
          throw new Error(data.error || 'Đã có lỗi xảy ra');
        }

        responseText = data.result;
        usedModel = data.model || selectedModel;
      }

      setResult(responseText);
    } catch (err: any) {
      console.error(err);
      setError(err.message || 'Đã có lỗi xảy ra khi tạo giáo án.');
    } finally {
      setIsLoading(false);
    }
  };

  const handleCopy = () => {
    if (result) {
      navigator.clipboard.writeText(result);
      setCopied(true);
      setTimeout(() => setCopied(false), 2000);
    }
  };

  return (
    <div className="min-h-screen bg-slate-50 text-slate-900 font-sans">
      <header className="bg-white border-b border-slate-200 sticky top-0 z-10 shadow-sm">
        <div className="max-w-6xl mx-auto px-4 h-16 flex items-center justify-between">
          <div className="flex items-center gap-3">
            <div className="w-10 h-10 bg-indigo-600 rounded-lg flex items-center justify-center text-white shrink-0">
              <BookOpen className="w-6 h-6" />
            </div>
            <div>
              <h1 className="text-xl font-bold tracking-tight text-slate-900">EduGen AI GA NLS - HuyHH</h1>
              <p className="text-xs text-slate-500 font-medium">Hệ thống Thiết kế Giáo án THCS CV5512</p>
            </div>
          </div>
          
          <div className="flex items-center gap-3">
            <a 
              href="https://aistudio.google.com/app/apikey" 
              target="_blank" 
              rel="noreferrer" 
              className="hidden sm:flex items-center gap-1.5 text-xs text-red-600 hover:text-red-700 font-bold hover:underline bg-red-50 px-3 py-1.5 rounded-lg border border-red-100 animate-pulse"
            >
              Lấy API key để sử dụng app
            </a>
            <button
              onClick={() => setShowSettings(true)}
              className={cn(
                "px-3.5 py-1.5 rounded-lg text-sm font-medium transition-all flex items-center gap-1.5 border shadow-sm",
                customApiKey 
                  ? "bg-white text-slate-700 border-slate-200 hover:bg-slate-50 hover:border-slate-300" 
                  : "bg-red-600 text-white border-red-700 hover:bg-red-700 hover:border-red-800 animate-bounce font-semibold"
              )}
            >
              <Settings className="w-4 h-4" />
              <span>Cấu hình AI</span>
              {!customApiKey && <span className="w-2 h-2 rounded-full bg-white inline-block animate-ping"></span>}
            </button>
          </div>
        </div>
      </header>

      <main className="max-w-6xl mx-auto px-4 py-8 grid grid-cols-1 lg:grid-cols-12 gap-8">
        
        {/* Left Column: Form */}
        <div className="lg:col-span-5 space-y-6">
          <div className="bg-white p-1 rounded-xl flex gap-1 border border-slate-200 shadow-sm">
            <button
              onClick={() => setMode('new')}
              className={cn(
                "flex-1 py-2.5 px-4 rounded-lg text-sm font-medium transition-all flex items-center justify-center gap-2",
                mode === 'new' ? "bg-indigo-50 text-indigo-700 shadow-sm" : "text-slate-600 hover:text-slate-900 hover:bg-slate-50"
              )}
            >
              <Sparkles className="w-4 h-4" />
              Khởi tạo Mới
            </button>
            <button
              onClick={() => setMode('upgrade')}
              className={cn(
                "flex-1 py-2.5 px-4 rounded-lg text-sm font-medium transition-all flex items-center justify-center gap-2",
                mode === 'upgrade' ? "bg-indigo-50 text-indigo-700 shadow-sm" : "text-slate-600 hover:text-slate-900 hover:bg-slate-50"
              )}
            >
              <FileText className="w-4 h-4" />
              Nâng cấp Giáo án
            </button>
          </div>

          <div className="bg-white rounded-2xl shadow-sm border border-slate-200 p-6 space-y-5">
            {mode === 'new' ? (
              <>
                <div>
                  <label className="block text-sm font-medium text-slate-700 mb-1.5">Tên bài học</label>
                  <input
                    type="text"
                    value={lessonName}
                    onChange={(e) => setLessonName(e.target.value)}
                    placeholder="VD: Định lý Thales trong tam giác"
                    className="w-full px-4 py-2.5 bg-slate-50 border border-slate-200 rounded-xl focus:outline-none focus:ring-2 focus:ring-indigo-500 focus:bg-white transition-colors text-sm"
                  />
                </div>
                
                <div className="grid grid-cols-2 gap-4">
                  <div>
                    <label className="block text-sm font-medium text-slate-700 mb-1.5">Môn học</label>
                    <input
                      type="text"
                      value={subject}
                      onChange={(e) => setSubject(e.target.value)}
                      placeholder="VD: Toán học"
                      className="w-full px-4 py-2.5 bg-slate-50 border border-slate-200 rounded-xl focus:outline-none focus:ring-2 focus:ring-indigo-500 focus:bg-white transition-colors text-sm"
                    />
                  </div>
                  <div>
                    <label className="block text-sm font-medium text-slate-700 mb-1.5">Lớp</label>
                    <input
                      type="text"
                      value={grade}
                      onChange={(e) => setGrade(e.target.value)}
                      placeholder="VD: 8"
                      className="w-full px-4 py-2.5 bg-slate-50 border border-slate-200 rounded-xl focus:outline-none focus:ring-2 focus:ring-indigo-500 focus:bg-white transition-colors text-sm"
                    />
                  </div>
                </div>

                <div>
                  <label className="block text-sm font-medium text-slate-700 mb-1.5">Thời lượng (Số tiết)</label>
                  <input
                    type="number"
                    value={duration}
                    onChange={(e) => setDuration(e.target.value)}
                    placeholder="VD: 2"
                    min="1"
                    className="w-full px-4 py-2.5 bg-slate-50 border border-slate-200 rounded-xl focus:outline-none focus:ring-2 focus:ring-indigo-500 focus:bg-white transition-colors text-sm"
                  />
                </div>
              </>
            ) : (
              <div className="space-y-4">
                <div>
                  <label className="block text-sm font-medium text-slate-700 mb-1.5 flex items-center justify-between">
                    <span>Tải tệp giáo án cũ (.docx, .pdf)</span>
                  </label>
                  
                  <div
                    onDragOver={handleDragOver}
                    onDragLeave={handleDragLeave}
                    onDrop={handleDrop}
                    className={cn(
                      "border-2 border-dashed rounded-xl p-5 text-center cursor-pointer transition-all duration-200 relative overflow-hidden flex flex-col items-center justify-center min-h-[140px]",
                      isDragging 
                        ? "border-indigo-500 bg-indigo-50/50 scale-[0.99]" 
                        : fileName
                          ? "border-emerald-300 bg-emerald-50/10 hover:border-emerald-400"
                          : "border-slate-200 bg-slate-50/50 hover:bg-slate-50 hover:border-slate-300"
                    )}
                    onClick={() => document.getElementById('file-upload')?.click()}
                  >
                    <input
                      id="file-upload"
                      type="file"
                      accept=".docx,.pdf"
                      onChange={handleFileChange}
                      className="hidden"
                    />
                    
                    {fileLoading ? (
                      <div className="flex flex-col items-center gap-2">
                        <Loader2 className="w-8 h-8 text-indigo-600 animate-spin" />
                        <span className="text-sm font-medium text-slate-600">Đang đọc và trích xuất dữ liệu từ file...</span>
                      </div>
                    ) : fileName ? (
                      <div className="flex flex-col items-center gap-2 w-full px-4">
                        <div className="w-12 h-12 bg-emerald-100 text-emerald-600 rounded-full flex items-center justify-center shadow-sm">
                          <Check className="w-6 h-6" />
                        </div>
                        <div className="text-sm font-semibold text-slate-700 truncate max-w-full">
                          {fileName}
                        </div>
                        <div className="text-xs text-slate-400">
                          Đã trích xuất thành công nội dung văn bản.
                        </div>
                        <button
                          type="button"
                          onClick={(e) => {
                            e.stopPropagation();
                            setFileName(null);
                            setContent('');
                          }}
                          className="mt-1 text-xs text-red-500 hover:text-red-700 font-medium hover:underline flex items-center gap-0.5"
                        >
                          <X className="w-3.5 h-3.5" /> Xóa file
                        </button>
                      </div>
                    ) : (
                      <div className="flex flex-col items-center gap-2">
                        <div className="w-12 h-12 bg-slate-100 text-slate-500 rounded-full flex items-center justify-center border border-slate-200 shadow-sm transition-colors">
                          <UploadCloud className="w-6 h-6 text-indigo-500" />
                        </div>
                        <div>
                          <span className="text-indigo-600 font-semibold text-sm hover:underline">Click để tải lên</span>
                          <span className="text-slate-500 text-sm"> hoặc kéo thả file vào đây</span>
                        </div>
                        <p className="text-xs text-slate-400">
                          Hỗ trợ định dạng Word (.docx) hoặc PDF (.pdf)
                        </p>
                      </div>
                    )}
                  </div>
                  
                  {fileError && (
                    <div className="mt-2 p-2 bg-red-50 text-red-600 text-xs rounded-lg border border-red-100 font-medium">
                      {fileError}
                    </div>
                  )}
                </div>

                <div>
                  <div className="flex items-center justify-between mb-1.5">
                    <label className="block text-sm font-medium text-slate-700">Nội dung giáo án cũ</label>
                    {content && (
                      <span className="text-[10px] text-slate-400 bg-slate-100 px-2 py-0.5 rounded font-mono">
                        {content.length.toLocaleString()} ký tự
                      </span>
                    )}
                  </div>
                  <textarea
                    value={content}
                    onChange={(e) => setContent(e.target.value)}
                    placeholder="Nội dung giáo án đã trích xuất sẽ hiển thị tại đây để bạn có thể xem và chỉnh sửa trực tiếp, hoặc dán trực tiếp văn bản giáo án..."
                    className="w-full px-4 py-3 bg-slate-50 border border-slate-200 rounded-xl focus:outline-none focus:ring-2 focus:ring-indigo-500 focus:bg-white transition-colors text-sm min-h-[220px] resize-y"
                  />
                </div>
              </div>
            )}

            <div className="pt-2">
              <button
                onClick={handleGenerate}
                disabled={isLoading}
                className="w-full bg-indigo-600 hover:bg-indigo-700 text-white font-medium py-3 px-4 rounded-xl shadow-sm transition-all flex items-center justify-center gap-2 disabled:opacity-70 disabled:cursor-not-allowed"
              >
                {isLoading ? (
                  <>
                    <Loader2 className="w-5 h-5 animate-spin" />
                    Đang thiết kế giáo án...
                  </>
                ) : (
                  <>
                    <Sparkles className="w-5 h-5" />
                    Tạo Giáo Án Chuẩn CV5512
                  </>
                )}
              </button>
            </div>
            
            {error && (
              <div className="p-3 bg-red-50 text-red-600 text-sm rounded-lg border border-red-100">
                {error}
              </div>
            )}
          </div>
        </div>

        {/* Right Column: Result */}
        <div className="lg:col-span-7 flex flex-col h-[calc(100vh-8rem)]">
          <div className="bg-white rounded-2xl shadow-sm border border-slate-200 flex-1 flex flex-col overflow-hidden">
            <div className="px-6 py-4 border-b border-slate-200 flex items-center justify-between shrink-0 bg-slate-50/50">
              <h2 className="font-semibold text-slate-800 flex items-center gap-2">
                <FileText className="w-4 h-4 text-indigo-500" />
                Kết quả Giáo án
              </h2>
              {result && (
                <div className="flex items-center gap-2">
                  <button
                    onClick={() => exportToDocx(result, lessonName || subject || "Giao_an_5512")}
                    className="text-sm font-medium text-slate-600 hover:text-indigo-600 flex items-center gap-1.5 transition-colors bg-white border border-slate-200 px-3 py-1.5 rounded-md shadow-sm"
                  >
                    <Download className="w-4 h-4 text-indigo-500" />
                    Tải file Word (.docx)
                  </button>
                  <button
                    onClick={handleCopy}
                    className="text-sm font-medium text-slate-600 hover:text-indigo-600 flex items-center gap-1.5 transition-colors bg-white border border-slate-200 px-3 py-1.5 rounded-md shadow-sm"
                  >
                    {copied ? <Check className="w-4 h-4 text-emerald-500" /> : <Copy className="w-4 h-4" />}
                    {copied ? 'Đã chép' : 'Chép nội dung'}
                  </button>
                </div>
              )}
            </div>
            
            <div className="flex-1 overflow-auto p-6 bg-white prose prose-slate max-w-none prose-headings:font-semibold prose-a:text-indigo-600 prose-table:border-collapse prose-th:border-slate-300 prose-th:bg-slate-100 prose-td:border-slate-300 prose-td:align-top">
              {result ? (
                <ReactMarkdown
                  remarkPlugins={[remarkGfm, remarkMath]}
                  rehypePlugins={[rehypeKatex]}
                >
                  {result}
                </ReactMarkdown>
              ) : (
                <div className="h-full flex flex-col items-center justify-center text-slate-400 space-y-4">
                  <div className="w-16 h-16 rounded-full bg-slate-50 flex items-center justify-center border border-slate-100 shadow-sm">
                    <Sparkles className="w-8 h-8 text-slate-300" />
                  </div>
                  <p className="text-sm max-w-sm text-center">
                    Điền thông tin và nhấn "Tạo Giáo Án" để nhận bản thiết kế chi tiết theo đúng chuẩn Công văn 5512.
                  </p>
                </div>
              )}
            </div>
          </div>
        </div>
      </main>

      {/* Settings Modal */}
      {showSettings && (
        <div className="fixed inset-0 bg-slate-900/60 backdrop-blur-sm flex items-center justify-center z-50 p-4">
          <div className="bg-white rounded-2xl shadow-2xl border border-slate-200 w-full max-w-md overflow-hidden animate-in fade-in zoom-in-95 duration-200">
            <div className="px-6 py-4 border-b border-slate-100 flex items-center justify-between bg-slate-50">
              <h3 className="font-semibold text-slate-900 flex items-center gap-2">
                <Settings className="w-4 h-4 text-indigo-500" />
                Cài đặt cấu hình AI
              </h3>
              {customApiKey && (
                <button 
                  onClick={() => setShowSettings(false)}
                  className="text-slate-400 hover:text-slate-600 rounded-lg p-1 hover:bg-slate-100 transition-colors"
                >
                  <X className="w-5 h-5" />
                </button>
              )}
            </div>
            
            <div className="p-6 space-y-6">
              {!customApiKey && (
                <div className="p-3 bg-red-50 text-red-700 text-xs rounded-lg border border-red-100 font-medium leading-relaxed">
                  Chào mừng bạn đến với **EduGen AI**! Ứng dụng yêu cầu API Key của bạn để kết nối trực tiếp với Gemini AI. Vui lòng cấu hình bên dưới để tiếp tục.
                </div>
              )}
              
              <div>
                <label className="block text-sm font-semibold text-slate-700 mb-1.5 flex items-center justify-between">
                  <span>Gemini API Key</span>
                  <a 
                    href="https://aistudio.google.com/app/apikey" 
                    target="_blank" 
                    rel="noreferrer"
                    className="text-xs text-indigo-600 hover:underline font-bold"
                  >
                    Lấy key miễn phí ở đây
                  </a>
                </label>
                <input
                  type="password"
                  placeholder="Nhập API Key của bạn (AI Studio)..."
                  defaultValue={customApiKey}
                  id="settings-api-key"
                  className="w-full px-4 py-2.5 bg-slate-50 border border-slate-200 rounded-xl focus:outline-none focus:ring-2 focus:ring-indigo-500 focus:bg-white transition-all text-sm font-mono shadow-sm"
                />
                <p className="mt-1.5 text-[10px] text-slate-400">
                  Key của bạn được lưu an toàn tại local storage, chúng tôi không lưu trữ trên máy chủ.
                </p>
              </div>
              
              <div>
                <label className="block text-sm font-semibold text-slate-700 mb-2">
                  Lựa chọn Model AI
                </label>
                {/* Model Cards Selector */}
                <div className="grid grid-cols-1 gap-2.5" id="settings-model-container">
                  {[
                    { id: 'gemini-2.5-flash', name: 'Gemini 2.5 Flash', desc: 'Tốc độ cực nhanh, thích hợp cho giáo án đơn giản.', tag: 'Khuyên dùng mặc định' },
                    { id: 'gemini-2.5-pro', name: 'Gemini 2.5 Pro', desc: 'Cực kỳ thông minh, tối ưu cho công thức toán & lồng ghép năng lực số phức tạp.', tag: 'Chất lượng cao nhất' },
                    { id: 'gemini-1.5-pro', name: 'Gemini 1.5 Pro', desc: 'Model thế hệ trước, độ ổn định cao.', tag: 'Truyền thống' }
                  ].map((model) => (
                    <button
                      key={model.id}
                      type="button"
                      onClick={() => {
                        const btnContainer = document.getElementById('settings-model-container');
                        btnContainer?.querySelectorAll('.model-card-btn').forEach(btn => {
                          btn.classList.remove('border-indigo-600', 'bg-indigo-50/30', 'ring-2', 'ring-indigo-500/20');
                          btn.classList.add('border-slate-200', 'bg-white');
                        });
                        const currentBtn = document.getElementById(`model-card-${model.id}`);
                        currentBtn?.classList.add('border-indigo-600', 'bg-indigo-50/30', 'ring-2', 'ring-indigo-500/20');
                        currentBtn?.classList.remove('border-slate-200', 'bg-white');
                        (window as any).tempSelectedModel = model.id;
                      }}
                      id={`model-card-${model.id}`}
                      className={cn(
                        "model-card-btn w-full p-3 text-left border rounded-xl transition-all duration-150 flex flex-col justify-between items-start gap-1 focus:outline-none",
                        (selectedModel === model.id)
                          ? "border-indigo-600 bg-indigo-50/30 ring-2 ring-indigo-500/20"
                          : "border-slate-200 bg-white hover:bg-slate-50"
                      )}
                    >
                      <div className="flex items-center justify-between w-full">
                        <span className="font-semibold text-slate-800 text-sm">{model.name}</span>
                        <span className={cn(
                          "text-[9px] px-1.5 py-0.5 rounded font-medium",
                          model.id === 'gemini-2.5-flash' ? "bg-indigo-100 text-indigo-700" :
                          model.id === 'gemini-2.5-pro' ? "bg-emerald-100 text-emerald-700" : "bg-slate-100 text-slate-600"
                        )}>
                          {model.tag}
                        </span>
                      </div>
                      <span className="text-[11px] text-slate-500 leading-normal">{model.desc}</span>
                    </button>
                  ))}
                </div>
              </div>
            </div>
            
            <div className="px-6 py-4 bg-slate-50 border-t border-slate-100 flex justify-end gap-2.5">
              {customApiKey && (
                <button
                  onClick={() => setShowSettings(false)}
                  className="px-4 py-2 text-sm font-medium text-slate-600 hover:text-slate-900 bg-white hover:bg-slate-100 border border-slate-200 rounded-xl transition-colors"
                >
                  Hủy
                </button>
              )}
              <button
                onClick={() => {
                  const keyVal = (document.getElementById('settings-api-key') as HTMLInputElement)?.value || '';
                  if (!keyVal.trim()) {
                    alert("Vui lòng nhập API Key để tiếp tục sử dụng ứng dụng.");
                    return;
                  }
                  const modelVal = (window as any).tempSelectedModel || selectedModel;
                  saveSettings(keyVal, modelVal);
                }}
                className="px-5 py-2 text-sm font-semibold text-white bg-indigo-600 hover:bg-indigo-700 rounded-xl shadow-sm transition-colors w-full sm:w-auto text-center"
              >
                Lưu và bắt đầu sử dụng
              </button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}
