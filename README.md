# Excel to Markdown & PDF Converter

Microsoft の MarkltDown ツールを活用して、Excel ファイルを Markdown と PDF に変換する Azure Function です。HTTP リクエストで簡単に起動して利用できます。

## 概要

このプロジェクトは Excel ファイルを次の形式に変換できる Azure Function を提供します：

- Excel → Markdown
- Excel → PDF
- Excel → Markdown + PDF

シンプルな実装で、Azure Functions 上に簡単にデプロイでき、Azure Blob Storage と連携して利用できます。

## シンプルなファイル構造

```
ExcelToMarkdown/
├── function_app.py      
├── requirements.txt     # 依存パッケージのリスト
├── host.json            # Azure Functions 設定
└── local.settings.json  # ローカル開発用設定
```

## 前提条件

- Azure アカウント
- Azure Blob Storage アカウント
- Python 3.10 以上
- Azure CLI または Azure Portal (デプロイ時)
- Azure Functions Core Tools (ローカル開発時)

## セットアップ手順

### 1. リポジトリのセットアップ

```bash
# フォルダを作成
mkdir ExcelToMarkdown
cd ExcelToMarkdown

# 各ファイルを作成
# function_app.py, requirements.txt, host.json, local.settings.json

# Python仮想環境を作成
python -m venv .venv

# 仮想環境をアクティブ化
# Windowsの場合:
.venv\Scripts\activate
# macOS/Linuxの場合:
source .venv/bin/activate

# 依存関係をインストール
pip install -r requirements.txt
```

### 2. ファイルの内容

#### function_app.py

```python
import azure.functions as func
import logging
import json
import base64
import tempfile
import os
import io
import pandas as pd
from datetime import datetime
from WeasyPrint import HTML
from markitdown import MarkItDown
from azure.storage.blob import BlobServiceClient

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
        
        if not req_body or 'inputContainer' not in req_body or 'inputBlobPath' not in req_body:
            return func.HttpResponse(
                json.dumps({"error": "リクエストに必要なパラメータが含まれていません"}),
                mimetype="application/json",
                status_code=400
            )
        
        # パラメータの取得
        input_container = req_body.get('inputContainer')
        input_blob_path = req_body.get('inputBlobPath')
        output_container = req_body.get('outputContainer', 'conversion-output')
        output_format = req_body.get('format', 'both').lower()  # 'markdown', 'pdf', または 'both'
        
        # 出力形式の検証
        if output_format not in ['markdown', 'pdf', 'both']:
            return func.HttpResponse(
                json.dumps({"error": "formatパラメータには 'markdown', 'pdf', または 'both' を指定してください"}),
                mimetype="application/json",
                status_code=400
            )
        
        # Blobストレージの接続文字列
        connect_str = os.environ["AzureWebJobsStorage"]
        
        # BlobServiceClientの作成
        blob_service_client = BlobServiceClient.from_connection_string(connect_str)
        
        # 入力コンテナクライアントの取得
        input_container_client = blob_service_client.get_container_client(input_container)
        
        # 入力Blobクライアントの取得
        input_blob_client = input_container_client.get_blob_client(input_blob_path)
        
        # Blobからデータをダウンロード
        blob_data = input_blob_client.download_blob()
        file_content = blob_data.readall()
        
        # ファイル名の取得
        file_name = os.path.basename(input_blob_path)
        
        # レスポンスデータの初期化
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        output_path_base = f"{timestamp}/{file_name.split('.')[0]}"
        
        response_data = {
            "status": "success",
            "message": f"Successfully processed {file_name}",
            "timestamp": timestamp,
            "files": {}
        }
        
        # 出力コンテナクライアントの取得（存在しない場合は作成）
        output_container_client = blob_service_client.get_container_client(output_container)
        if not output_container_client.exists():
            output_container_client.create_container()
        
        # 形式に応じて変換処理を実行
        images = []
        
        if output_format in ['markdown', 'both']:
            # Markdownへの変換
            try:
                file_stream = io.BytesIO(file_content)
                converter = MarkItDown()
                md_result = converter.convert_stream(file_stream, filename=file_name)
                
                # Markdown内容をBlobにアップロード
                md_file_name = f"{file_name.split('.')[0]}.md"
                md_blob_path = f"{output_path_base}.md"
                md_blob_client = output_container_client.get_blob_client(md_blob_path)
                md_blob_client.upload_blob(md_result.text_content, overwrite=True)
                
                # 画像ファイルの処理
                if md_result.images:
                    for img_idx, img in enumerate(md_result.images):
                        img_name = f"image{img_idx+1}.png"
                        img_path = f"{timestamp}/media/{img_name}"
                        img_blob_client = output_container_client.get_blob_client(img_path)
                        img_blob_client.upload_blob(img.content, overwrite=True)
                        
                        images.append({
                            "name": img_name,
                            "path": img_path,
                            "container": output_container,
                            "url": img_blob_client.url,
                            "contentType": "image/png",
                            "size": len(img.content)
                        })
                
                response_data["files"]["markdown"] = {
                    "name": md_file_name,
                    "path": md_blob_path,
                    "container": output_container,
                    "url": md_blob_client.url
                }
                
            except Exception as e:
                return func.HttpResponse(
                    json.dumps({
                        "status": "error",
                        "message": f"Markdown変換エラー: {str(e)}"
                    }),
                    mimetype="application/json",
                    status_code=500
                )
        
        if output_format in ['pdf', 'both']:
            # PDFへの変換
            try:
                handle, temp_path = tempfile.mkstemp(suffix='.xlsx')
                try:
                    with os.fdopen(handle, 'wb') as temp_file:
                        temp_file.write(file_content)
                    
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
                    
                    # PDFコンテンツをBlobにアップロード
                    pdf_file_name = f"{file_name.split('.')[0]}.pdf"
                    pdf_blob_path = f"{output_path_base}.pdf"
                    pdf_blob_client = output_container_client.get_blob_client(pdf_blob_path)
                    pdf_blob_client.upload_blob(pdf_content, overwrite=True)
                    
                    response_data["files"]["pdf"] = {
                        "name": pdf_file_name,
                        "path": pdf_blob_path,
                        "container": output_container,
                        "url": pdf_blob_client.url
                    }
                    
                finally:
                    # 一時ファイルの削除
                    try:
                        os.unlink(temp_path)
                    except:
                        pass
            except Exception as e:
                return func.HttpResponse(
                    json.dumps({
                        "status": "error",
                        "message": f"PDF変換エラー: {str(e)}"
                    }),
                    mimetype="application/json",
                    status_code=500
                )
        
        # 画像配列を追加
        if images:
            response_data["files"]["images"] = images
        
        return func.HttpResponse(
            body=json.dumps(response_data),
            mimetype="application/json",
            status_code=200
        )
    
    except Exception as e:
        logging.error(f"リクエスト処理中にエラーが発生しました: {str(e)}")
        return func.HttpResponse(
            json.dumps({
                "status": "error",
                "message": f"内部サーバーエラー: {str(e)}"
            }),
            mimetype="application/json",
            status_code=500
        )
```

#### requirements.txt

```
azure-functions>=1.14.0
markitdown[all]>=0.1.0
WeasyPrint>=55.0
pandas>=1.3.0
openpyxl>=3.0.10
azure-storage-blob>=12.14.0
```

#### host.json

```json
{
  "version": "2.0",
  "logging": {
    "applicationInsights": {
      "samplingSettings": {
        "isEnabled": true,
        "excludedTypes": "Request"
      }
    }
  },
  "extensionBundle": {
    "id": "Microsoft.Azure.Functions.ExtensionBundle",
    "version": "[3.*, 4.0.0)"
  },
  "functionTimeout": "00:05:00"
}
```

#### local.settings.json

```json
{
  "IsEncrypted": false,
  "Values": {
    "AzureWebJobsStorage": "DefaultEndpointsProtocol=https;AccountName=yourstorageaccount;AccountKey=yourstoragekey;EndpointSuffix=core.windows.net",
    "FUNCTIONS_WORKER_RUNTIME": "python"
  }
}
```

## ローカル開発と実行

### 1. Azure Functions Core Tools のインストール

```bash
# NPM経由でインストール
npm install -g azure-functions-core-tools@4 --unsafe-perm true
```

### 2. ローカルでの実行

```bash
# プロジェクトディレクトリで
func start
```

これにより、ローカル環境（通常は http://localhost:7071）で関数が起動します。

## Azure Functions へのデプロイ手順

### 方法1: Azure CLIを使用したデプロイ

#### 1. Azureにログイン

```bash
az login
```

#### 2. リソースグループの作成

```bash
az group create --name ExcelConverterGroup --location japaneast
```

#### 3. ストレージアカウントの作成

```bash
az storage account create --name excelconvstorage --location japaneast --resource-group ExcelConverterGroup --sku Standard_LRS
```

#### 4. Function Appの作成

```bash
az functionapp create --resource-group ExcelConverterGroup --consumption-plan-location japaneast --runtime python --runtime-version 3.10 --functions-version 4 --name ExcelConverterApp --storage-account excelconvstorage --os-type linux
```

#### 5. アプリケーションのデプロイ

```bash
func azure functionapp publish ExcelConverterApp
```

### 方法2: Azure Portalを使用したデプロイ

#### 1. Azure Portalにログイン

[Azure Portal](https://portal.azure.com)にアクセスしてログインします。

#### 2. Function Appの作成

1. 「リソースの作成」をクリック
2. 「Function App」を検索して選択
3. 次の設定を行います：
   - サブスクリプション: あなたのサブスクリプション
   - リソースグループ: 新規作成または既存のものを選択
   - 関数アプリ名: 一意の名前を指定（例: ExcelConverterApp）
   - 公開: コード
   - ランタイムスタック: Python
   - バージョン: 3.10
   - 地域: 東日本
4. 「確認および作成」をクリックし、検証が完了したら「作成」をクリック

#### 3. コードのデプロイ

1. Visual Studio Code に Azure Functions 拡張機能をインストール
2. Azure にサインイン
3. プロジェクトを開き、F1キーを押して「Azure Functions: Deploy to Function App」を選択
4. 作成した Function App を選択
5. デプロイが完了するのを待つ

## 使用方法

### 1. Azure Blob Storage の準備

1. Azure Portal でストレージアカウントを作成または既存のものを使用
2. `excel-input` などの名前でコンテナを作成（入力用）
3. 必要に応じて `conversion-output` などの名前でコンテナを作成（出力用）
4. Excel ファイルを入力用コンテナにアップロード

### 2. Function App へのリクエスト

Function App に対して以下のような JSON データで POST リクエストを送信します：

```json
{
  "inputContainer": "excel-input",
  "inputBlobPath": "SAMPLE_会議室予約システム機能設計書.xlsx",
  "outputContainer": "conversion-output",
  "format": "both"  // "markdown", "pdf", "both" から選択可能
}
```

#### リクエストパラメータ

- `inputContainer`: 入力ファイルが格納されている Azure Blob Storage のコンテナ名（必須）
- `inputBlobPath`: 変換するExcelファイルのパス（必須）
- `outputContainer`: 変換結果を保存する Azure Blob Storage のコンテナ名（オプション、デフォルトは "conversion-output"）
- `format`: 出力形式、以下のいずれか（オプション、デフォルトは "both"）
  - `"markdown"`: Markdownのみに変換
  - `"pdf"`: PDFのみに変換
  - `"both"`: MarkdownとPDFの両方に変換

### 3. 変換結果の保存場所

変換結果は以下の場所に保存されます：

- Markdownファイル： `conversion-output/[実行日時]/[元ファイル名].md`
- PDFファイル： `conversion-output/[実行日時]/[元ファイル名].pdf`
- 画像ファイル： `conversion-output/[実行日時]/media/[画像ファイル名]`

### 4. レスポンス形式

```json
{
  "status": "success",
  "message": "Successfully processed example.xlsx",
  "timestamp": "20250123_123456",
  "files": {
    "markdown": {
      "name": "example.md",
      "path": "20250123_123456/example.md",
      "container": "conversion-output",
      "url": "https://..."
    },
    "pdf": {
      "name": "example.pdf",
      "path": "20250123_123456/example.pdf",
      "container": "conversion-output",
      "url": "https://..."
    },
    "images": [
      {
        "name": "image1.png",
        "path": "20250123_123456/media/image1.png",
        "container": "conversion-output",
        "url": "https://...",
        "contentType": "image/png",
        "size": 12345
      }
    ]
  }
}
```

## トラブルシューティング

### 1. デプロイ時のエラー

- 依存関係のインストールに問題がある場合は、`requirements.txt`の内容を確認してください
- Linux環境でWeasyPrintのインストールに問題がある場合は、追加のシステムライブラリが必要かもしれません

### 2. 実行時のエラー

- ファイルサイズが大きすぎる場合は、Azure FunctionsのデフォルトのHTTPリクエストサイズ制限に注意してください
- タイムアウトが発生する場合は、`host.json`の`functionTimeout`を調整してください
- ストレージへの接続エラーが発生した場合は、接続文字列を確認してください

### 3. Azure Functionsのログを確認

```bash
# ログのストリーミング表示
func azure functionapp logstream ExcelConverterApp
```

## 参考リソース

- [Microsoft MarkltDown](https://github.com/microsoft/markitdown) - Markdownへの変換に使用
- [Azure Functions Python開発者ガイド](https://docs.microsoft.com/azure/azure-functions/functions-reference-python)
- [Azure Blob Storage の使用](https://docs.microsoft.com/azure/storage/blobs/storage-quickstart-blobs-python)
- [WeasyPrint](https://weasyprint.org/) - HTML/CSSからPDFへの変換に使用