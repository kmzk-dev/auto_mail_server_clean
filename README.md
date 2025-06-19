# MailManager Python Class
`MailManager`は、IMAPサーバー上のメールをプログラムで管理し、自動化するためのPythonクラスです。メールの接続、メールIDの取得、ヘッダー情報の抽出、削除フラグの設定、メールの完全削除、フォルダ一覧の取得といった基本的な機能を備えています。

## Overviews
このクラスは、以下のような特定のニーズを持つ組織や個人を支援するために作成されました。

- **営業メールの追跡と集計:** 営業担当者がBCCで追跡用メールアドレスを設定して送信したメールを自動的に集計し、営業活動の進捗を把握します。
- **マーケティングメールのレスポンス集計:** 特定のマーケティングキャンペーンに対するメールのレスポンス（例：フォーム送信後の自動返信メールなど）を集計します。
- **メールサーバーのクリーニング:** 集計が完了したログメールや不要なメールを定期的にサーバーから削除し、ディスク容量の節約とサーバーパフォーマンスの維持に貢献します。

これらの用途では、メールの本文内容ではなく、主にヘッダー情報（送信元、送信先、件名、日時など）に基づく「カウント」が重要であり、集計後はメールを「破棄」する必要があるという特性があります。

## Features
- IMAP4/IMAP4_SSL プロトコルをサポート
- IMAPサーバーへのセキュアな接続と認証
- 指定したフォルダ内の全メールIDの取得
- 各メールのヘッダー情報（日時、送信元、送信先、件名）の抽出とデコード
- メールへの削除フラグ付与と完全削除（expunge）
- サーバー上の全メールボックス（フォルダ）の一覧取得

## Install
特別なインストールは不要です。Pythonの標準ライブラリ imaplib, email, re, email.header, email.utils を使用しています。

## How To Use
`MailManager`クラスを使用するには、IMAPサーバーの情報（`ホスト名`,`ポート`,`アカウント`,`パスワード`）をインスタンス化時に提供します。

使用例:
```python
from mail_manager_function import MailManager # クラスが別ファイルの場合

# IMAPサーバーの接続情報
IMAP_SERVER = "your.imap.server.com"
IMAP_PORT = 993 # SSL/TLSの場合
IMAP_ACCOUNT = "your_email@example.com"
IMAP_PASSWORD = "your_email_password"

# MailManagerのインスタンス化
mail_manager = MailManager(IMAP_SERVER, IMAP_PORT, IMAP_ACCOUNT, IMAP_PASSWORD)

try:
    # 接続
    mail_manager.connect()

    # フォルダ一覧の取得
    folders = mail_manager.GetFolders()

    # "Inbox" フォルダのメールIDを取得
    mail_manager.getIds(directory="Inbox")

    # 取得したメールのヘッダー情報をフェッチ
    headers = mail_manager.fetchHeaders()
    for header in headers:
        print(header)

    # ここでヘッダー情報に基づいて集計やフィルタリングのロジックを実装
    # 例: 特定の送信元からのメールをカウント
    # count_from_specific_sender = sum(1 for row in headers[1:] if "specific_sender@example.com" in row[1])
    # print(f"特定の送信元からのメール数: {count_from_specific_sender}")

    # 集計後、メールに削除フラグを付与（注意：実行するとメールが削除されます）
    mail_manager.markDeleteFlag()

    # 削除フラグが付いたメールを完全に削除（注意：元に戻せません）
    mail_manager.Delete()

except Exception as e:
    print(f"処理中にエラーが発生しました: {e}")
finally:
    # 接続を切断
    if mail_manager.mail_connection:
        mail_manager.disconnect()
```

## Others
この MailManager クラスは、IMAP4プロトコルに準拠しており、Python標準の imaplib を使用しているため、一般的なIMAPサーバーに幅広く対応できます。SSL/TLSによるセキュアな接続もサポートしています。

一般的なIMAPサーバーであれば、このクラスをそのまま利用できる可能性が高いです。

## Caution
- **非標準的なIMAP実装:** ごく稀に、特定のメールプロバイダがIMAPプロトコルを非標準的な方法で実装している場合、予期しない動作をする可能性があります。
- **特定の認証方式:** 現在はユーザー名とパスワードによる認証を前提としています。OAuth2.0など、より高度な認証方式を要求するサーバーには直接対応していません。
- **サーバー固有のフォルダ名:** メールボックス（フォルダ）の名前はサーバーによって異なる場合があります（例: "Inbox", "受信トレイ", "INBOX"など）。GetFolders() メソッドで利用可能なフォルダ名を確認し、必要に応じて getIds() メソッドの directory 引数を調整してください。
- **レートリミット:** 短時間に大量の操作（特に fetchHeaders などで多数のメールを連続で取得する場合）を行うと、サーバー側で一時的に接続が制限される（レートリミット）可能性があります。
- **メール本文の取得:** 現在、このクラスはメールのヘッダー情報のみを取得します。メール本文（テキスト、HTML、添付ファイルなど）の取得機能は含まれていません。本文の内容に基づいて処理を行いたい場合は、機能拡張が必要です。
- **エラーハンドリング:** 基本的なエラーは出力されますが、本番運用する際には、より堅牢なエラーハンドリング（リトライ処理、特定の例外に対する詳細なメッセージなど）を実装することが推奨されます。

## License
MIT License