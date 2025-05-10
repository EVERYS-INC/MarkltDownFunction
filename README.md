# Excel to Markdown & PDF Converter

Microsoft の MarkltDown ツールを活用して、Excel ファイルを Markdown および PDF に変換する Azure Function 実装です。

## 概要

このプロジェクトは Excel ファイルを以下の形式に変換できる Azure Function を提供します：

- Excel → Markdown
- Excel → PDF
- Excel → Markdown + PDF

Azure Blob Storage に保存された Excel ファイルを処理し、変換結果を同じく Azure Blob Storage に保存します。

## 必要条件

- Azure アカウント
- Azure CLI (バージョン 2.40.0 以上)
- Git

## デプロイ手順

### 1. リポジトリのクローン

```bash
git clone https://github.com/EVERYS-INC/MarkltDownFunction.git
cd MarkltDownFunction
```

### 2. Azure リソースのセットアップ

```bash
# 変数設定（自分の名前を小文字で設定）
NAME=<your-name>
myResourceGroup=<your-resource-group-name>
LOCATION=japaneast

# 既存のストレージアカウントを利用する場合は下記を設定。違うリソースグループにあるストレージアカウントを指定すると
ConnectionString=<your-storage-connection-string>

# ストレージアカウント作成 ※既存のストレージアカウントを利用する場合はスキップ
az storage account create --name mystorageaccount${NAME} \
                          --resource-group ${myResourceGroup} \
                          --location ${LOCATION} \
                          --sku Standard_LRS

# ストレージアカウントの接続文字列を取得 ※既存のストレージアカウントを利用する場合はスキップ
ConnectionString=$(az storage account show-connection-string \
                  --name mystorageaccount${NAME} \
                  --resource-group ${myResourceGroup} \
                  --query connectionString \
                  --output tsv)

# Function App作成
az functionapp create \
  --name exceltomarkdownapp-${NAME} \
  --resource-group ${myResourceGroup} \
  --storage-account mystorageaccount${NAME} \
  --consumption-plan-location ${LOCATION} \
  --runtime python \
  --runtime-version 3.10 \
  --functions-version 4 \
  --os-type linux

# Blobストレージのコンテナ作成
az storage container create --name excel-input \
  --connection-string "${ConnectionString}"

az storage container create --name conversion-output \
  --connection-string "${ConnectionString}"

# Function Appにアプリケーション設定を追加
az functionapp config appsettings set --name exceltomarkdownapp-${NAME} \
                                      --resource-group ${myResourceGroup} \
                                      --settings STORAGE_CONNECTION_STRING="${ConnectionString}"

# Function Appにデプロイ
az functionapp deployment source config-zip --name exceltomarkdownapp-${NAME} \
                                           --resource-group ${myResourceGroup} \
                                           --src $(zip -r /tmp/function.zip . > /dev/null && echo /tmp/function.zip)
```

### 3. デプロイの確認

デプロイが成功したら、Azure Portal で Function App の詳細を確認できます。

```bash
# 関数のURLを取得
FUNCTION_URL=$(az functionapp function show \
  --name exceltomarkdownapp-${NAME} \
  --resource-group ${myResourceGroup} \
  --function-name convert \
  --query invokeUrlTemplate \
  --output tsv)

# 関数キーを取得
FUNCTION_KEY=$(az functionapp function keys list \
  --name exceltomarkdownapp-${NAME} \
  --resource-group ${myResourceGroup} \
  --function-name convert \
  --query default \
  --output tsv)

echo "Function URL: $FUNCTION_URL"
echo "Function Key: $FUNCTION_KEY"
```

## 使用方法

1. Excelファイルを `excel-input` コンテナにアップロード

```bash
# サンプルExcelファイルをアップロード
az storage blob upload \
  --container-name excel-input \
  --name "設計書.xlsx" \
  --file "./sample/設計書.xlsx" \
  --connection-string "${ConnectionString}"
```

2. Function App に下記JSONをPOSTリクエスト

```json
{
  "inputContainer": "excel-input",
  "inputBlobPath": "設計書.xlsx",
  "outputContainer": "conversion-output",
  "format": "both"  // "markdown", "pdf", "both" から選択可能
}
```

3. 変換結果は以下の場所に保存

- Markdownファイル： `conversion-output/[実行日時]/[元ファイル名].md`
- PDFファイル： `conversion-output/[実行日時]/[元ファイル名].pdf`
- 画像ファイル： `conversion-output/[実行日時]/media/[画像ファイル名]`

## レスポンス形式

```json
{
  "status": "success",
  "message": "Successfully processed 設計書.xlsx",
  "timestamp": "20250123_123456",
  "files": {
    "markdown": {
      "name": "設計書.md",
      "path": "20250123_123456/設計書.md",
      "container": "conversion-output",
      "url": "https://..."
    },
    "pdf": {
      "name": "設計書.pdf",
      "path": "20250123_123456/設計書.pdf",
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

## 注意事項

- 大きなExcelファイルを処理する場合は、Function App のタイムアウト設定を確認してください
- WeasyPrintはLinux環境での依存関係があるため、Function App は Linux ベースを使用しています
- Function App の実行プランによっては、メモリやCPUの制限がある場合があります

## 参考リソース

- [Microsoft MarkltDown](https://github.com/microsoft/markitdown)
- [Azure Functions の Python 開発者向けリファレンス](https://docs.microsoft.com/ja-jp/azure/azure-functions/functions-reference-python)