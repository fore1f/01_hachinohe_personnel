from fpdf import FPDF
import os

def create_manual():
    pdf = FPDF()
    pdf.add_page()
    
    # 日本語フォント（MSゴシック）の追加
    font_path = r"C:\Windows\Fonts\msgothic.ttc"
    if not os.path.exists(font_path):
        font_path = r"C:\Windows\Fonts\msmincho.ttc"

    pdf.add_font("msgothic", "", font_path, uni=True)
    pdf.set_font("msgothic", size=16)
    
    # タイトル
    pdf.cell(0, 10, txt="人事異動データ抽出ツール 操作マニュアル", ln=True, align='C')
    pdf.ln(2)
    
    # 1. プログラム一覧
    pdf.set_font("msgothic", size=13)
    pdf.cell(0, 8, txt="1. プログラム一覧と役割", ln=True)
    pdf.set_font("msgothic", size=11)
    # multi_cellの引数を調整し、不自然な空白を防ぐ
    pdf.multi_cell(0, 7, txt="各プログラムはそれぞれの組織専用の抽出ロジック（シート名や列配置の判定）を持っています。")
    pdf.ln(1)
    
    # 一覧テーブル
    pdf.set_fill_color(240, 240, 240)
    pdf.cell(50, 7, txt="組織名", border=1, fill=True)
    pdf.cell(50, 7, txt="ファイル名", border=1, fill=True)
    pdf.cell(45, 7, txt="uv (推奨)", border=1, fill=True)
    pdf.cell(45, 7, txt="python", border=1, fill=True, ln=True)
    
    # 各行（高さを7に統一）
    data = [
        ("八戸市（市長部局）", "hachinohesi.py", "uv run hachinohesi.py", "python hachinohesi.py"),
        ("広域事務組合", "hachi_kouiki.py", "uv run hachi_kouiki.py", "python hachi_kouiki.py"),
        ("消防本部", "hachi_shoubou.py", "uv run hachi_shoubou.py", "python hachi_shoubou.py"),
        ("水道企業団", "hachi_suidou.py", "uv run hachi_suidou.py", "python hachi_suidou.py"),
    ]
    
    pdf.set_font("msgothic", size=10)
    for row in data:
        pdf.cell(50, 7, txt=row[0], border=1)
        pdf.cell(50, 7, txt=row[1], border=1)
        pdf.cell(45, 7, txt=row[2], border=1)
        pdf.cell(45, 7, txt=row[3], border=1, ln=True)
    pdf.ln(4)
    
    # 2. 使い方
    pdf.set_font("msgothic", size=13)
    pdf.cell(0, 8, txt="2. 基本的な使いかた", ln=True)
    pdf.set_font("msgothic", size=11)
    # テキスト内の不要な位置での改行やスペースを削除し、multi_cellの自動整形に任せる
    usage_text = (
        "1. コマンドの実行: プログラムがあるフォルダをターミナルで開き、実行コマンドを入力します。\n"
        "   ※uvがある環境では「uv run ～」、それ以外では「python ～」を使用してください。\n"
        "2. ファイルの選択: エクセルの選択画面が開きますので、対象のファイルを選択してください。\n"
        "3. 自動抽出: プログラムがシートを解析し、データを抽出します。\n"
        "4. 結果の確認: 同じ場所に「～_抽出結果.txt」という名前のファイルが生成されます。"
    )
    pdf.multi_cell(0, 7, txt=usage_text)
    pdf.ln(4)
    
    # 3. 体裁ルール
    pdf.set_font("msgothic", size=13)
    pdf.cell(0, 8, txt="3. 出力テキストのルール", ln=True)
    pdf.set_font("msgothic", size=11)
    rules_text = (
        "・インデント: 各データ行の先頭にはタブが入り、右側に寄せて表示されます。\n"
        "・項目区切り: 職名、旧職、氏名の間はタブで区切られています。\n"
        "・括弧なし: 旧職や採用などの前後にあった（ ）は削除されています。\n"
        "・記号なし: 見出しの ■ 記号などは削除されています。"
    )
    pdf.multi_cell(0, 7, txt=rules_text)
    pdf.ln(4)
    
    # 4. 警告機能
    pdf.set_font("msgothic", size=13)
    pdf.cell(0, 8, txt="4. 姓名区切りチェック機能", ln=True)
    pdf.set_font("msgothic", size=11)
    warn_text = (
        "消防本部のプログラム等では、氏名の中に全角スペースがない場合、自動で検出します。\n"
        "処理の最後に、区切り漏れの可能性がある氏名をまとめて表示しますので、確認・修正を行ってください。"
    )
    pdf.multi_cell(0, 7, txt=warn_text)
    
    # 保存
    output_filename = "人事異動抽出ツール_操作マニュアル.pdf"
    try:
        pdf.output(output_filename)
        print(f"PDFが作成されました: {output_filename}")
    except PermissionError:
        print(f"エラー: {output_filename} が開かれているため保存できません。閉じてから再実行してください。")

if __name__ == "__main__":
    create_manual()
