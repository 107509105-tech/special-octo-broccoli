"""
方案 2: Unstructured
使用 Unstructured 庫進行智能文檔解析
自動識別文檔結構（標題、段落、表格等）
"""

from unstructured.partition.pdf import partition_pdf
from unstructured.chunking.title import chunk_by_title
from typing import List, Dict, Any
import json
from pathlib import Path


class UnstructuredPDFParser:
    def __init__(self, pdf_path: str):
        self.pdf_path = pdf_path
        self.elements = []
        self.chunks = []

    def parse(self, strategy: str = "hi_res") -> List[Any]:
        """
        解析 PDF 文檔

        參數:
            strategy: 解析策略
                - "auto": 自動選擇
                - "fast": 快速模式（僅文本）
                - "hi_res": 高精度模式（包含 OCR，處理圖像）
        """
        print(f"開始解析 PDF: {self.pdf_path}")
        print(f"使用策略: {strategy}")

        try:
            self.elements = partition_pdf(
                filename=self.pdf_path,
                strategy=strategy,
                infer_table_structure=True,  # 推斷表格結構
                extract_images_in_pdf=False,  # 不提取圖像（加快速度）
                extract_image_block_types=["Image", "Table"],
                extract_image_block_to_payload=False,
                languages=["chi_tra", "eng"],  # 繁體中文和英文
            )

            print(f"成功解析，提取了 {len(self.elements)} 個元素")
            return self.elements

        except Exception as e:
            print(f"解析時發生錯誤: {e}")
            print("嘗試使用 fast 策略...")

            # 降級到 fast 策略
            self.elements = partition_pdf(
                filename=self.pdf_path,
                strategy="fast"
            )
            print(f"使用 fast 策略成功解析，提取了 {len(self.elements)} 個元素")
            return self.elements

    def analyze_elements(self) -> Dict[str, int]:
        """分析文檔元素類型"""
        element_types = {}

        for element in self.elements:
            element_type = type(element).__name__
            element_types[element_type] = element_types.get(element_type, 0) + 1

        return element_types

    def convert_to_documents(self) -> List[Dict[str, Any]]:
        """將元素轉換為文檔格式"""
        documents = []

        for idx, element in enumerate(self.elements):
            doc = {
                "id": idx,
                "text": str(element),
                "type": type(element).__name__,
                "metadata": element.metadata.to_dict() if hasattr(element, 'metadata') else {}
            }
            documents.append(doc)

        return documents

    def chunk_documents(self,
                       max_characters: int = 1000,
                       new_after_n_chars: int = 800,
                       overlap: int = 200) -> List[Dict[str, Any]]:
        """
        使用智能分塊策略
        根據文檔標題結構進行分塊
        """
        print("\n開始智能分塊...")

        try:
            # 使用標題分塊
            chunked_elements = chunk_by_title(
                elements=self.elements,
                max_characters=max_characters,
                new_after_n_chars=new_after_n_chars,
                overlap=overlap,
                overlap_all=True
            )

            self.chunks = []
            for idx, chunk in enumerate(chunked_elements):
                chunk_doc = {
                    "chunk_id": idx,
                    "text": str(chunk),
                    "type": type(chunk).__name__,
                    "metadata": {
                        **(chunk.metadata.to_dict() if hasattr(chunk, 'metadata') else {}),
                        "chunk_index": idx,
                        "source": self.pdf_path
                    }
                }
                self.chunks.append(chunk_doc)

            print(f"生成了 {len(self.chunks)} 個智能塊")
            return self.chunks

        except Exception as e:
            print(f"智能分塊失敗: {e}")
            print("使用簡單分塊策略...")

            # 降級到簡單分塊
            return self._simple_chunk(max_characters, overlap)

    def _simple_chunk(self, chunk_size: int, overlap: int) -> List[Dict[str, Any]]:
        """簡單的滑動窗口分塊"""
        documents = self.convert_to_documents()
        chunks = []

        current_chunk = ""
        current_metadata = {}
        chunk_id = 0

        for doc in documents:
            text = doc["text"]

            if len(current_chunk) + len(text) <= chunk_size:
                current_chunk += text + "\n"
                current_metadata.update(doc["metadata"])
            else:
                if current_chunk:
                    chunks.append({
                        "chunk_id": chunk_id,
                        "text": current_chunk.strip(),
                        "metadata": {
                            **current_metadata,
                            "chunk_index": chunk_id,
                            "source": self.pdf_path
                        }
                    })
                    chunk_id += 1

                current_chunk = text + "\n"
                current_metadata = doc["metadata"].copy()

        # 添加最後一個塊
        if current_chunk:
            chunks.append({
                "chunk_id": chunk_id,
                "text": current_chunk.strip(),
                "metadata": {
                    **current_metadata,
                    "chunk_index": chunk_id,
                    "source": self.pdf_path
                }
            })

        self.chunks = chunks
        return chunks

    def save_elements(self, output_path: str):
        """保存原始元素"""
        documents = self.convert_to_documents()
        with open(output_path, 'w', encoding='utf-8') as f:
            json.dump(documents, f, ensure_ascii=False, indent=2)
        print(f"原始元素已保存到: {output_path}")

    def save_chunks(self, output_path: str):
        """保存分塊結果"""
        with open(output_path, 'w', encoding='utf-8') as f:
            json.dump(self.chunks, f, ensure_ascii=False, indent=2)
        print(f"分塊結果已保存到: {output_path}")


def main():
    # 設定 PDF 路徑
    pdf_path = "ysm20r.pdf"

    # 創建解析器
    parser = UnstructuredPDFParser(pdf_path)

    # 解析 PDF（使用 fast 策略以加快速度，如需更高精度可改為 "hi_res"）
    elements = parser.parse(strategy="fast")

    # 分析元素類型
    print("\n=== 文檔元素類型統計 ===")
    element_types = parser.analyze_elements()
    for elem_type, count in sorted(element_types.items(), key=lambda x: x[1], reverse=True):
        print(f"{elem_type}: {count}")

    # 保存原始元素
    print("\n保存原始解析結果...")
    parser.save_elements("unstructured_elements.json")

    # 智能分塊
    print("\n準備 RAG 文檔塊...")
    chunks = parser.chunk_documents(
        max_characters=1000,
        new_after_n_chars=800,
        overlap=200
    )

    # 保存分塊結果
    parser.save_chunks("unstructured_chunks.json")

    # 顯示統計信息
    print("\n=== 解析統計 ===")
    print(f"總元素數: {len(elements)}")
    print(f"RAG 塊數: {len(chunks)}")

    # 顯示第一個塊的示例
    if chunks:
        print("\n=== 第一個 RAG 塊示例 ===")
        print(f"文本長度: {len(chunks[0]['text'])}")
        print(f"類型: {chunks[0].get('type', 'N/A')}")
        print(f"元數據: {chunks[0]['metadata']}")
        print(f"文本預覽: {chunks[0]['text'][:200]}...")

    # 顯示不同類型元素的示例
    print("\n=== 元素類型示例 ===")
    shown_types = set()
    for element in elements[:50]:  # 只檢查前50個元素
        elem_type = type(element).__name__
        if elem_type not in shown_types:
            shown_types.add(elem_type)
            print(f"\n{elem_type}:")
            print(f"  {str(element)[:100]}...")
            if len(shown_types) >= 5:  # 只顯示前5種類型
                break


if __name__ == "__main__":
    main()
