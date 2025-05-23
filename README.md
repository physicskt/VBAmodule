# VBAmodule

## Overview / 概要

This project is a VBA module manager for Excel workbooks that allows users to easily import and export VBA modules. It provides utilities for managing VBA code across different Excel files, enabling better code organization and reuse.

このプロジェクトは、Excel ワークブック用の VBA モジュールマネージャーであり、ユーザーが VBA モジュールを簡単にインポートおよびエクスポートできるようにするものです。異なる Excel ファイル間で VBA コードを管理するためのユーティリティを提供し、コードの整理と再利用を促進します。

## Features / 機能

- Export all VBA modules from a workbook to a specified folder
- Import VBA modules from a folder into a workbook
- Support for different module types (standard modules, class modules, form modules)
- Prevention of duplicate imports
- User-friendly dialog for selecting import folders
- Error handling for import/export operations

- ワークブックからすべての VBA モジュールを指定したフォルダにエクスポート
- フォルダから VBA モジュールをワークブックにインポート
- 異なるモジュールタイプ（標準モジュール、クラスモジュール、フォームモジュール）のサポート
- 重複インポートの防止
- インポートフォルダを選択するためのユーザーフレンドリーなダイアログ
- インポート/エクスポート操作のエラー処理

## Installation / インストール方法

There are two ways to install and use the VBA module manager:

1. Use the included `macro.xlsm` file which already contains the BasManager module.
2. Import the `BasManager.bas` file from the ExportedModules folder into your own Excel workbook.

To import the module into your own workbook:
1. Open your workbook in Excel
2. Press Alt+F11 to open the VBA editor
3. Right-click on your project in the Project Explorer
4. Select Import File and navigate to the BasManager.bas file

VBA モジュールマネージャーをインストールして使用するには、次の 2 つの方法があります：

1. BasManager モジュールがすでに含まれている `macro.xlsm` ファイルを使用する。
2. ExportedModules フォルダから `BasManager.bas` ファイルを自分の Excel ワークブックにインポートする。

モジュールを自分のワークブックにインポートするには：
1. Excel でワークブックを開く
2. Alt+F11 を押して VBA エディタを開く
3. プロジェクトエクスプローラーでプロジェクトを右クリック
4. 「ファイルのインポート」を選択し、BasManager.bas ファイルに移動する

## Usage / 使用方法

### Exporting Modules / モジュールのエクスポート

To export all VBA modules from your workbook:

1. Run the `DoExportAllModules()` subroutine
2. Modules will be exported to an "ExportedModules" folder in the same directory as your workbook
3. A message will appear when the export is complete

ワークブックからすべての VBA モジュールをエクスポートするには：

1. `DoExportAllModules()` サブルーチンを実行する
2. モジュールはワークブックと同じディレクトリにある「ExportedModules」フォルダにエクスポートされる
3. エクスポートが完了すると、メッセージが表示される

### Importing Modules / モジュールのインポート

To import VBA modules into your workbook:

1. Run the `DoImportAllModules()` subroutine
2. A dialog will appear for you to select the folder containing the modules
3. The modules will be imported into your workbook (existing modules with the same name will not be overwritten)
4. A message will appear when the import is complete

VBA モジュールをワークブックにインポートするには：

1. `DoImportAllModules()` サブルーチンを実行する
2. モジュールを含むフォルダを選択するためのダイアログが表示される
3. モジュールがワークブックにインポートされる（同じ名前の既存モジュールは上書きされない）
4. インポートが完了すると、メッセージが表示される

## Contribution / 貢献方法

Contributions to this project are welcome! Here's how you can contribute:

1. Fork the repository
2. Create a new branch for your feature or bugfix
3. Make your changes
4. Submit a pull request

Please ensure your code follows the existing style and includes appropriate comments.

このプロジェクトへの貢献を歓迎します！貢献する方法は次のとおりです：

1. リポジトリをフォークする
2. 機能やバグ修正のための新しいブランチを作成する
3. 変更を加える
4. プルリクエストを送信する

コードが既存のスタイルに従っており、適切なコメントが含まれていることを確認してください。

## License / ライセンス

This project is licensed under the MIT License - see the LICENSE file for details.

このプロジェクトは MIT ライセンスの下でライセンスされています - 詳細については LICENSE ファイルを参照してください。