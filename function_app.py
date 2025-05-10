import azure.functions as func
import logging
import json
import base64
import tempfile
import os
import io
import pandas as pd
from WeasyPrint import HTML
from markitdown import MarkItDown

# Azure Function App定義
app = func.FunctionApp()

@app.route(route="convert", auth_level=func.AuthLevel.FUNCTION)
async def excel_converter(req: func.HttpRequest) -> func.HttpResponse:
    """
    Excel ファイルを Markdown と PDF に変換する Azure Function
    """
    logging.info('Excel変換処理を開始します')
    
    try:
        # リクエストボディのJSONを解析
        req_body = req.get_json()
        
        if not req_body or 'file' not in req_body:
            return func.HttpResponse(
                json.dumps({"error": "base64エンコードされたファイルをリクエストボディに含めてください"}),
                mimetype="application/json",
                status_code=400
            )
        
        # ファイルコンテンツとパラメータの取得
        file_content = req_body.get('file')
        file_name = req_body.get('filename', 'document')
        output_format = req_body.get('format', 'both').lower()  # 'markdown', 'pdf', または 'both'
        
        # 出力形式の検証
        if output_format not in ['markdown', 'pdf', 'both']:
            return func.HttpResponse(
                json.dumps({"error": "formatパラメータには 'markdown', 'pdf', または 'both' を指定してください"}),
                mimetype="application/json",
                status_code=400
            )
        
        # base64デコード
        try:
            decoded_content = base64.b64decode(file_content)
        except Exception as e:
            return func.HttpResponse(
                json.dumps({"error": f"base64デコードに失敗しました: {str(e)}"}),
                mimetype="application/json",
                status_code=400
            )
        
        # レスポンスデータの初期化
        response_data = {
            "original_filename": file_name
        }
        
        # 形式に応じて変換処理を実行
        if output_format in ['markdown', 'both']:
            # Markdownへの変換
            try:
                file_stream = io.BytesIO(decoded_content)
                converter = MarkItDown()
                md_result = converter.convert_stream(file_stream, filename=file_name)
                
                response_data.update({
                    "markdown_success": True,
                    "markdown_title": md_result.title,
                    "markdown_content": md_result.text_content
                })
            except Exception as e:
                response_data.update({
                    "markdown_success": False,
                    "markdown_error": f"Markdown変換エラー: {str(e)}"
                })
        
        if output_format in ['pdf', 'both']:
            # PDFへの変換
            try:
                handle, temp_path = tempfile.mkstemp(suffix='.xlsx')
                try:
                    with os.fdopen(handle, 'wb') as temp_file:
                        temp_file.write(decoded_content)
                    
                    # Excelファイルの読み込みと変換
                    xl = pd.ExcelFile(temp_path)
                    sheet_names = xl.sheet_names
                    
                    # HTMLを生成
                    html_content = "<html><head><style>"
                    html_content += """
                    body { font-family: Arial, sans-serif; }
                    h1 { color: #0066cc; }
                    table { border-collapse: collapse; width: 100%; margin-bottom: 20px; }
                    th, td { border: 1px solid #ddd; padding: 8px; text-align: left; }
                    th { background-color: #f2f2f2; }
                    .sheet-title { margin-top: 30px; margin-bottom: 10px; }
                    """
                    html_content += "</style></head><body>"
                    html_content += f"<h1>{file_name}</h1>"
                    
                    # 各シートの処理
                    for sheet_name in sheet_names:
                        df = pd.read_excel(temp_path, sheet_name=sheet_name)
                        html_content += f"<h2 class='sheet-title'>Sheet: {sheet_name}</h2>"
                        html_content += df.to_html(index=False)
                    
                    html_content += "</body></html>"
                    
                    # HTMLからPDFへの変換
                    pdf_buffer = io.BytesIO()
                    HTML(string=html_content).write_pdf(pdf_buffer)
                    pdf_buffer.seek(0)
                    pdf_content = pdf_buffer.read()
                    
                    # base64エンコード
                    pdf_base64 = base64.b64encode(pdf_content).decode('utf-8')
                    
                    response_data.update({
                        "pdf_success": True,
                        "pdf_content": pdf_base64
                    })
                    
                finally:
                    # 一時ファイルの削除
                    try:
                        os.unlink(temp_path)
                    except:
                        pass
            except Exception as e:
                response_data.update({
                    "pdf_success": False,
                    "pdf_error": f"PDF変換エラー: {str(e)}"
                })
        
        return func.HttpResponse(
            body=json.dumps(response_data),
            mimetype="application/json",
            status_code=200
        )
    
    except Exception as e:
        logging.error(f"リクエスト処理中にエラーが発生しました: {str(e)}")
        return func.HttpResponse(
            json.dumps({"error": f"内部サーバーエラー: {str(e)}"}),
            mimetype="application/json",
            status_code=500
        )