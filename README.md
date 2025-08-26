# 📄 Extração de Dados de PDFs com OCR e Exportação para Excel
 
## 🧾 Descrição
 
Este script automatiza a extração de números destacados em arquivos PDF utilizando OCR (Reconhecimento Óptico de Caracteres) com Tesseract. Ele gera imagens de uma área específica da primeira página de cada PDF, extrai os números contidos nessa área e os salva em um arquivo Excel.
 
## 📂 Estrutura de Pastas
 
- `pdfs/`: Pasta onde os arquivos PDF devem ser colocados.
- `prints/`: Pasta onde as imagens geradas dos PDFs serão salvas temporariamente.
- `Resumo_Faturamento VALE_MAI25.xlsx`: Arquivo Excel onde os dados extraídos serão armazenados.
 
## 🛠️ Requisitos
 
- Python 3.x
- Bibliotecas:
  - `PyMuPDF` (`fitz`)
  - `Pillow`
  - `pytesseract`
  - `openpyxl`
 
Instale as dependências com:
 
```bash
pip install pymupdf pillow pytesseract openpyxl
```
 
Além disso, é necessário ter o **Tesseract OCR** instalado no sistema. O caminho para o executável deve ser configurado corretamente no script:
 
```python
pytesseract.pytesseract.tesseract_cmd = r'C:\Users\Vinicius.cunha\AppData\Local\Programs\Tesseract-OCR\tesseract.exe'
```
 
## ⚙️ Funcionamento
 
1. **Leitura dos PDFs**: O script percorre todos os arquivos `.pdf` na pasta `pdfs`.
2. **Geração de Imagem**: Para cada PDF, é gerada uma imagem de uma área específica da primeira página.
3. **Extração de Números**: A imagem é processada com OCR para extrair números.
4. **Seleção de Índices**: Apenas os números nos índices `[0, 1, 3, 4, 5, 7]` são selecionados.
5. **Exportação para Excel**: Os números extraídos são inseridos em colunas específicas do Excel, junto com o nome do PDF.
 
## 📌 Observações
 
- O script salva os dados a partir da linha 12 do Excel, procurando a próxima linha vazia na coluna 9.
- As imagens geradas são salvas na pasta `prints`, mas podem ser removidas após o processamento.
- O script está preparado para lidar com erros e continuar o processamento dos demais arquivos.
 
## ✅ Resultado
 
Ao final, você terá um arquivo Excel preenchido com os dados extraídos dos PDFs, organizados por nome de arquivo e colunas correspondentes aos números destacados
