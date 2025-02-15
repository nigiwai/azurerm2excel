# azurerm2excel

## 説明

azure環境を構築したtfstateファイルからパラメタシート（Excel）を作成します

## 前提条件

- Python 3.11以上
- pip (Pythonパッケージインストーラー)

## 準備

1. 仮想環境を作成してアクティブにします:

    ```sh
    python -m venv venv
    .\venv\Scripts\activate.ps1

    ```

2. 必要なパッケージをインストールします:

    python 3.11.9で確認

    ```sh
    pip install -r requirements.txt
    ```

## 使用方法

```cmd
python azurerm2excel.py <path_to_tfstate_file> <description_folder1> [<description_folder2> ...]
python .\azurerm2excel.py .\tfstate\terraform.tfstate D:\git\terra2excel\json\azurerm_4.14.0 D:\git\terra2excel\json\azuread_3.1.0
```

## ライセンス

MITライセンス
