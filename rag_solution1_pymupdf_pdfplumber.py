"""
方案 1: PyMuPDF + pdfplumber
結合 PyMuPDF 的速度優勢和 pdfplumber 的表格提取能力
"""

import fitz  # PyMuPDF
import pdfplumber
import json
from typing import List, Dict, Any
from pathlib import Path


class PDFParser:
    def __init__(self, pdf_path: str):
        self.pdf_path = pdf_path
        self.documents = []

    def extract_text_with_pymupdf(self) -> List[Dict[str, Any]]:
        """使用 PyMuPDF 快速提取文本"""
        doc = fitz.open(self.pdf_path)
        text_documents = []

        for page_num in range(len(doc)):
            page = doc[page_num]
            text = page.get_text()

            # 提取元數據
            text_documents.append({
                "page": page_num + 1,
                "content": text,
                "content_type": "text",
                "metadata": {
                    "source": self.pdf_path,
                    "page": page_num + 1,
                    "total_pages": len(doc)
                }
            })

        doc.close()
        return text_documents

    def extract_tables_with_pdfplumber(self) -> List[Dict[str, Any]]:
        """使用 pdfplumber 提取表格"""
        table_documents = []

        with pdfplumber.open(self.pdf_path) as pdf:
            for page_num, page in enumerate(pdf.pages):
                tables = page.extract_tables()

                if tables:
                    for table_idx, table in enumerate(tables):
                        # 將表格轉換為文本格式
                        table_text = self._table_to_text(table)

                        table_documents.append({
                            "page": page_num + 1,
                            "content": table_text,
                            "content_type": "table",
                            "table_data": table,
                            "metadata": {
                                "source": self.pdf_path,
                                "page": page_num + 1,
                                "table_index": table_idx,
                                "doc_type": "technical_specification"
                            }
                        })

        return table_documents

    def _table_to_text(self, table: List[List[str]]) -> str:
        """將表格轉換為易於檢索的文本格式"""
        if not table:
            return ""

        text_lines = []
        # 假設第一行是表頭
        headers = table[0]

        for row in table[1:]:
            row_text = []
            for header, cell in zip(headers, row):
                if header and cell:
                    row_text.append(f"{header}: {cell}")
            if row_text:
                text_lines.append(" | ".join(row_text))

        return "\n".join(text_lines)

    def parse(self) -> List[Dict[str, Any]]:
        """執行完整的解析流程"""
        print(f"開始解析 PDF: {self.pdf_path}")

        # 提取文本
        print("步驟 1: 使用 PyMuPDF 提取文本...")
        text_docs = self.extract_text_with_pymupdf()
        print(f"  提取了 {len(text_docs)} 頁文本")

        # 提取表格
        print("步驟 2: 使用 pdfplumber 提取表格...")
        table_docs = self.extract_tables_with_pdfplumber()
        print(f"  提取了 {len(table_docs)} 個表格")

        # 合併所有文檔
        self.documents = text_docs + table_docs

        return self.documents

    def save_to_json(self, output_path: str):
        """保存解析結果為 JSON"""
        with open(output_path, 'w', encoding='utf-8') as f:
            json.dump(self.documents, f, ensure_ascii=False, indent=2)
        print(f"結果已保存到: {output_path}")

    def prepare_for_rag(self, chunk_size: int = 1000, overlap: int = 200) -> List[Dict[str, Any]]:
        """準備用於 RAG 的文檔塊"""
        chunks = []

        for doc in self.documents:
            content = doc["content"]

            # 對長文本進行分塊
            if len(content) > chunk_size:
                # 簡單的滑動窗口分塊
                start = 0
                while start < len(content):
                    end = start + chunk_size
                    chunk_text = content[start:end]

                    chunks.append({
                        "text": chunk_text,
                        "metadata": {
                            **doc["metadata"],
                            "content_type": doc["content_type"],
                            "chunk_start": start,
                            "chunk_end": end
                        }
                    })

                    start += chunk_size - overlap
            else:
                chunks.append({
                    "text": content,
                    "metadata": {
                        **doc["metadata"],
                        "content_type": doc["content_type"]
                    }
                })

        return chunks


def main():
    # 設定 PDF 路徑
    pdf_path = "ysm20r.pdf"

    # 創建解析器
    parser = PDFParser(pdf_path)

    # 解析 PDF
    documents = parser.parse()

    # 保存原始解析結果
    parser.save_to_json("parsed_documents.json")

    # 準備 RAG 分塊
    print("\n步驟 3: 準備 RAG 文檔塊...")
    rag_chunks = parser.prepare_for_rag(chunk_size=1000, overlap=200)
    print(f"  生成了 {len(rag_chunks)} 個文檔塊")

    # 保存 RAG 分塊
    with open("rag_chunks.json", 'w', encoding='utf-8') as f:
        json.dump(rag_chunks, f, ensure_ascii=False, indent=2)
    print(f"RAG 分塊已保存到: rag_chunks.json")

    # 顯示統計信息
    print("\n=== 解析統計 ===")
    print(f"總文檔數: {len(documents)}")
    print(f"文本頁數: {len([d for d in documents if d['content_type'] == 'text'])}")
    print(f"表格數量: {len([d for d in documents if d['content_type'] == 'table'])}")
    print(f"RAG 塊數: {len(rag_chunks)}")

    # 顯示第一個文檔塊的示例
    if rag_chunks:
        print("\n=== 第一個 RAG 塊示例 ===")
        print(f"文本長度: {len(rag_chunks[0]['text'])}")
        print(f"元數據: {rag_chunks[0]['metadata']}")
        print(f"文本預覽: {rag_chunks[0]['text'][:200]}...")


if __name__ == "__main__":
    main()
