import os
import shutil
import yaml
import re
from PyPDF2 import PdfReader
import logging
from pathlib import Path
import hashlib

# ロギングの設定
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

def extract_title_and_author(pdf_path):
    """PDFから論文のタイトルと最初の著者を抽出する"""
    try:
        reader = PdfReader(pdf_path)
        
        # PDF全体のテキストを取得（最初の2ページのみ）
        text = ""
        max_pages = min(2, len(reader.pages))
        for i in range(max_pages):
            page_text = reader.pages[i].extract_text()
            if page_text:
                text += page_text + "\n"
        
        if not text:
            logger.warning(f"PDFからテキストを抽出できませんでした: {pdf_path}")
            return os.path.splitext(os.path.basename(pdf_path))[0], "Unknown"
        
        # タイトルと副題を分割して抽出
        full_title, subtitle = extract_title_and_subtitle(text, reader.metadata)
        
        # 著者の抽出 - 複数の方法を試みる
        author = extract_author(text, full_title, os.path.basename(pdf_path), reader.metadata)
        
        # 不要な文字や記号を除去、著者名の整形
        full_title = re.sub(r'[\n\r\t]+', ' ', full_title).strip()
        author = re.sub(r'[\n\r\t]+', ' ', author).strip()
        
        # 長すぎる著者名は切り詰める
        if len(author) > 50:
            author = author[:47] + "..."
        
        logger.info(f"抽出結果 - タイトル: {full_title}")
        logger.info(f"抽出結果 - 著者: {author}")
        
        return full_title, author
        
    except Exception as e:
        logger.error(f"PDFの処理中にエラーが発生しました: {pdf_path} - {str(e)}")
        # エラーが発生した場合はファイル名をタイトルとして、著者は不明とする
        base_name = os.path.splitext(os.path.basename(pdf_path))[0]
        return base_name, "Unknown"

def extract_title_and_subtitle(text, metadata=None):
    """PDFからタイトルと副題を抽出する"""
    # タイトルの抽出方法 - 複数の方法を試みる
    title = None
    subtitle = None
    
    # 方法1: メタデータからの抽出（参考情報だが信頼性は低い）
    metadata_title = None
    if metadata and metadata.get('/Title'):
        metadata_title = metadata.get('/Title', '').strip()
        if len(metadata_title) > 10:  # 最低限の長さがあるタイトルのみ
            # メタデータのタイトルがコロンで区切られている場合は主題と副題に分割
            if ':' in metadata_title:
                parts = metadata_title.split(':', 1)
                title = parts[0].strip()
                subtitle = parts[1].strip() if len(parts) > 1 else None
            else:
                title = metadata_title
    
    # 方法2: 一般的な論文パターンからの抽出
    if not title:
        # 最初の数行を分析
        lines = text.split('\n')[:30]  # 最初の30行を対象
        
        # 空行を除去
        lines = [line.strip() for line in lines if line.strip()]
        
        # タイトルになりそうな行を特定
        title_candidates = []
        for line in lines:
            # 明らかに著者や所属などの行を除外
            if re.search(r'(University|Institute|Department|Abstract|Introduction|Keywords|©|Email|http)', line, re.IGNORECASE):
                continue
            # 通常著者名は短いので、ある程度の長さがあるものをタイトル候補とする
            if 20 <= len(line) <= 200 and not line.startswith('Fig') and not re.search(r'^[0-9]', line):
                title_candidates.append(line)
        
        if title_candidates:
            # タイトルが主題と副題に分かれているか確認（コロンや改行で区切られている場合）
            main_candidate = title_candidates[0]
            if ':' in main_candidate:
                parts = main_candidate.split(':', 1)
                title = parts[0].strip()
                subtitle = parts[1].strip() if len(parts) > 1 else None
            else:
                title = main_candidate
                
                # 次の行が副題の可能性がある
                if len(title_candidates) > 1 and len(title_candidates[1]) < len(title) * 1.5:
                    second_line = title_candidates[1]
                    # 副題の特徴: 主題より短く、"for", "of", "in", "on", "with", "using"などで始まることが多い
                    if re.match(r'^(for|of|in|on|with|using|a|an|the|toward)', second_line.lower()):
                        subtitle = second_line
    
    # 方法3: 特定のパターンマッチング
    if not title:
        # タイトルのパターン
        title_patterns = [
            r'(?:Title|TITLE)[:\s]+(.*?)[\n\r]',
            r'^([A-Z][^.!?]*[.!?])(?:\s|$)',  # 文頭から最初のピリオドまで
            r'^\s*([A-Z][^.!?]{10,100}[.!?])'  # 十分な長さの文
        ]
        
        for pattern in title_patterns:
            match = re.search(pattern, text)
            if match:
                matched_text = match.group(1).strip()
                # コロンで区切られている場合は主題と副題に分割
                if ':' in matched_text:
                    parts = matched_text.split(':', 1)
                    title = parts[0].strip()
                    subtitle = parts[1].strip() if len(parts) > 1 else None
                else:
                    title = matched_text
                break
    
    # タイトルが見つからない場合はファイル名を使用
    if not title:
        title = os.path.splitext(os.path.basename(pdf_path))[0]
    
    # タイトルが異常に長い場合は切り詰める
    if title and len(title) > 150:
        title = title[:147] + "..."
    
    # 副題があれば主題に含める（括弧で囲む）
    full_title = title
    if subtitle:
        # 副題が異常に長い場合は切り詰める
        if len(subtitle) > 100:
            subtitle = subtitle[:97] + "..."
        # 副題が全て大文字の場合は、先頭だけ大文字にする（読みやすさのため）
        if subtitle.isupper() and len(subtitle) > 10:
            subtitle = subtitle.capitalize()
        full_title = f"{title}: {subtitle}"
    
    return full_title, subtitle

def extract_author(text, title, filename, metadata=None):
    """PDFから著者情報を抽出する"""
    author = None
    
    # 方法1: メタデータからの抽出 (信頼性が低いので補助的に使用)
    metadata_author = None
    if metadata and metadata.get('/Author'):
        metadata_author = metadata.get('/Author', '').strip()
        if metadata_author and len(metadata_author) > 2 and len(metadata_author) < 50:
            # 複数著者の場合は最初の著者のみを取得
            if ',' in metadata_author:
                author = metadata_author.split(',')[0].strip()
            elif ';' in metadata_author:
                author = metadata_author.split(';')[0].strip()
            elif 'and' in metadata_author.lower():
                author = metadata_author.split('and')[0].strip()
            else:
                author = metadata_author
    
    # 方法2: ArXiv識別子からの著者情報抽出（arxivの論文の場合）
    if not author:
        arxiv_id = None
        # ファイル名からarxiv IDを検出
        arxiv_match = re.search(r'(\d{4}\.\d{5})(v\d+)?', filename)
        if arxiv_match:
            arxiv_id = arxiv_match.group(1)
            logger.info(f"ArXiv IDを検出: {arxiv_id}")
            
            # ページ内容からの著者検索（arxivの論文は特定のフォーマットを持つ）
            lines = text.split('\n')[:50]  # 最初の50行を対象
            
            # arxivの論文では通常タイトルの下に著者リストがある
            title_index = -1
            for i, line in enumerate(lines):
                if title and title.strip() in line.strip():
                    title_index = i
                    break
            
            if title_index >= 0 and title_index + 1 < len(lines):
                # タイトル直後の行を著者行として処理
                potential_authors = lines[title_index + 1]
                # 著者行の検証（一般的な著者リストの特徴）
                if validate_author_line(potential_authors):
                    # 一般的な著者リストの整理（カンマ、セミコロン、「and」で区切られている）
                    if ',' in potential_authors:
                        author = potential_authors.split(',')[0].strip()
                    elif ';' in potential_authors:
                        author = potential_authors.split(';')[0].strip()
                    elif 'and' in potential_authors.lower():
                        author = potential_authors.split('and')[0].strip()
                    else:
                        # 単一著者または区切りがない場合
                        author = potential_authors.strip()
                    
                    # 著者名の整形（括弧内の所属情報などを削除）
                    if author:
                        author = re.sub(r'\(.*?\)', '', author).strip()
                        # 電子メールアドレスを削除
                        author = re.sub(r'\S+@\S+', '', author).strip()
                
                # 著者名の検証（通常、著者名は短く、数字を含まない）
                if author and not validate_author_name(author):
                    author = None  # 無効な著者名をリセット
    
    # 方法3: タイトル後のテキストからの著者検索
    if not author:
        # タイトル後の短いテキスト部分から著者を探す
        title_index = -1
        if title in text:
            title_index = text.find(title) + len(title)
        
        if title_index > 0:
            author_text = text[title_index:title_index + 1000]  # より広い範囲で検索
            
            # まず数行を抽出
            author_lines = author_text.split('\n')[:10]
            
            # 著者パターンをチェック
            for i, line in enumerate(author_lines):
                if validate_author_line(line):
                    # 名前らしきパターンを探す
                    name_match = re.search(r'([A-Z][a-z]+(?:\s+[A-Z][a-z]+)+)', line)
                    if name_match:
                        author = name_match.group(1).strip()
                        # 著者名を検証
                        if validate_author_name(author):
                            break
                    
                    # カンマで区切られた名前を探す
                    if ',' in line and not re.search(r'\d', line.split(',')[0]):
                        author = line.split(',')[0].strip()
                        # 著者名を検証
                        if validate_author_name(author):
                            break
            
            # 名前らしきパターンを正規表現で検索
            if not author:
                author_patterns = [
                    r'(?:Author|AUTHORS|By)[:\s]+(.*?)[\n\r]',
                    r'(?:\n|\r)((?:[A-Z][a-z]+\s+)+[A-Z][a-z]+)(?:\n|\r|,)',  # 名前のパターン (例: John Smith)
                    r'([A-Z][a-z]+\s+[A-Z][a-z]+)(?:\s*,\s*\d|\s*\(\d)',  # 名前の後にカンマか括弧付き数字
                    r'([A-Z][a-z]+\s+[A-Z]\.\s*[A-Z][a-z]+)',  # ミドルネームイニシャル付き
                    r'([A-Z][a-z]+(?:-[A-Z][a-z]+)?\s+[A-Z][a-z]+)'  # ハイフン付き名前も対応
                ]
                
                for pattern in author_patterns:
                    match = re.search(pattern, author_text)
                    if match:
                        potential_author = match.group(1).strip()
                        # 著者名を検証
                        if validate_author_name(potential_author):
                            author = potential_author
                            break
            
            # 著者情報が見つかった場合、さらに加工
            if author:
                # 複数著者の場合は最初の著者のみを取得
                if ',' in author:
                    author = author.split(',')[0].strip()
                elif ';' in author:
                    author = author.split(';')[0].strip()
                elif ' and ' in author.lower():
                    author = author.split(' and ')[0].strip()
                
                # メールアドレスやその他の不要情報を削除
                author = re.sub(r'\S+@\S+', '', author).strip()
                author = re.sub(r'\(.*?\)', '', author).strip()
    
    # 方法4: 論文全体からの著者検索（最後の手段）
    if not author:
        # 学術論文で頻出する著者表記パターン
        author_global_patterns = [
            r'([A-Z][a-z]+\s+[A-Z]\.\s*[A-Z][a-z]+)',  # ミドルネームイニシャル付き
            r'([A-Z][a-z]+\s+[A-Z][a-z]+)(?:\s*,\s*\d|\s*\(\d)',  # 名前の後にカンマか括弧付き数字
            r'([A-Z][a-z]+\s+(?:[A-Z][a-z]+\s+){0,2}[A-Z][a-z]+)(?=\s*,|\s*and|\s*;|\s*\n)'  # 複数の単語で構成される名前
        ]
        
        for pattern in author_global_patterns:
            match = re.search(pattern, text[:3000])  # 最初の3000文字だけ対象
            if match:
                potential_author = match.group(1).strip()
                # 著者名を検証
                if validate_author_name(potential_author):
                    author = potential_author
                    break
    
    # 著者が見つからない場合は "Unknown" を使用
    if not author or len(author) < 3:
        # metadata_authorがある場合はそれを使用
        if metadata_author and len(metadata_author) > 2:
            author = metadata_author
        else:
            author = "Unknown"
    
    return author

def validate_author_line(line):
    """著者行らしいかどうかを検証する"""
    line = line.strip()
    
    # 著者行の特徴
    # 1. 一般的に短い（100文字未満）
    if len(line) > 100:
        return False
    
    # 2. 「Abstract」「Introduction」などの論文セクション見出しではない
    if re.search(r'(abstract|keywords|introduction|university|institute|^fig|table|copyright)', line.lower()):
        return False
    
    # 3. 通常は数字から始まらない
    if re.match(r'^\d', line):
        return False
    
    # 4. 通常は大文字小文字が混在する
    if line.isupper() and len(line) > 5:
        return False
    
    # 5. FOR、OF、IN などで始まる場合は副題である可能性が高い
    if re.match(r'^(FOR|OF|IN|ON|WITH|USING|A|AN|THE|TOWARD)\s', line.upper()):
        return False
    
    return True

def validate_author_name(name):
    """抽出された著者名が有効かどうかを検証する"""
    name = name.strip()
    
    # 著者名の検証ルール
    # 1. 長さが適切（2〜50文字）
    if len(name) < 2 or len(name) > 50:
        return False
    
    # 2. 通常、著者名には少なくとも1つの空白がある（名と姓）
    if ' ' not in name:
        return False
    
    # 3. 通常、著者名は数字を含まない
    if re.search(r'\d', name):
        return False
    
    # 4. 通常、著者名は大文字で始まる単語（一部の単語が大文字で始まる）
    # イニシャル（例：J. K. Rowling）もOK
    if not re.search(r'[A-Z][a-z]+|[A-Z]\.', name):
        return False
    
    # 5. 通常、著者名が非常に長いフレーズの場合は無効
    words = name.split()
    if len(words) > 6:  # 名前、ミドルネーム、姓の場合でも通常6語以下
        return False
    
    # 6. 特定のキーワードを含む場合は著者名ではない可能性が高い
    keywords = ['university', 'institute', 'department', 'abstract', 'introduction', 
                'keywords', 'copyright', 'rights', 'reserved', 'published', 
                'submitted', 'received', 'accepted', 'revised']
    for keyword in keywords:
        if keyword in name.lower():
            return False
    
    # 7. 一般的な前置詞や接続詞だけで構成されていないこと
    prepositions = ['for', 'of', 'in', 'on', 'with', 'using', 'by', 'to', 'at', 'from', 'and']
    if all(word.lower() in prepositions for word in words):
        return False
    
    # 8. 大文字のみの単語が含まれる場合は副題の可能性が高い
    if any(word.isupper() and len(word) > 1 for word in words):
        if not re.match(r'[A-Z]\.', name):  # イニシャルは例外
            return False
    
    return True

def sanitize_filename(filename):
    """ファイル名に使用できない文字を置換する"""
    # Windowsでファイル名に使用できない文字を置換
    invalid_chars = r'[<>:"/\\|?*]'
    sanitized = re.sub(invalid_chars, '_', filename)
    # ファイル名の長さ制限（Windowsは255文字）
    if len(sanitized) > 200:  # 余裕を持って200文字に制限
        sanitized = sanitized[:197] + "..."
    return sanitized

def main():
    try:
        # YAML設定ファイルの読み込み
        config_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "config.yaml")
        if not os.path.exists(config_path):
            logger.error(f"設定ファイルが見つかりません: {config_path}")
            # サンプル設定ファイルを作成
            sample_config = {
                "input_folders": ["./papers"],  # 論文PDFが保存されているフォルダ
                "output_folder": "./outputs",  # 出力先フォルダ
                "processed_folder": "./processed_papers"  # 処理済みファイルの移動先フォルダ
            }
            with open(config_path, 'w', encoding='utf-8') as f:
                yaml.dump(sample_config, f, default_flow_style=False, allow_unicode=True)
            logger.info(f"サンプル設定ファイルを作成しました: {config_path}")
            return
        
        with open(config_path, 'r', encoding='utf-8') as f:
            config = yaml.safe_load(f)
        
        # 複数の入力フォルダに対応
        input_folders = config.get('input_folders', [])
        # 後方互換性のため、古い形式のinput_folderもサポート
        if not input_folders and 'input_folder' in config:
            input_folders = [config.get('input_folder')]
            
        output_folder = config.get('output_folder', './outputs')
        processed_folder = config.get('processed_folder', './processed_papers')
        
        # 入力フォルダが指定されていない場合はエラー
        if not input_folders:
            logger.error("入力フォルダが設定されていません。")
            return
        
        # 相対パスを絶対パスに変換（出力フォルダ）
        if not os.path.isabs(output_folder):
            output_folder = os.path.join(os.path.dirname(os.path.abspath(__file__)), output_folder)
        
        # 相対パスを絶対パスに変換（処理済みフォルダ）
        if not os.path.isabs(processed_folder):
            processed_folder = os.path.join(os.path.dirname(os.path.abspath(__file__)), processed_folder)
        
        # 出力フォルダと処理済みフォルダが存在しない場合は作成
        os.makedirs(output_folder, exist_ok=True)
        os.makedirs(processed_folder, exist_ok=True)
        
        # 処理状況のカウント
        total_files = 0
        processed_files = 0
        skipped_files = 0
        moved_files = 0
        
        # 処理済みファイルのハッシュを記録
        processed_file_hashes = set()
        
        # 各入力フォルダを処理
        for folder_path in input_folders:
            # 相対パスを絶対パスに変換（入力フォルダ）
            if not os.path.isabs(folder_path):
                folder_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), folder_path)
            
            if not os.path.exists(folder_path):
                logger.warning(f"指定された入力フォルダが存在しません: {folder_path}")
                continue
                
            logger.info(f"フォルダを処理中: {folder_path}")
            
            # 入力フォルダ内のPDFファイルを再帰的に検索
            for root, _, files in os.walk(folder_path):
                for file in files:
                    if file.lower().endswith('.pdf'):
                        total_files += 1
                        pdf_path = os.path.join(root, file)
                        
                        # ファイルハッシュを計算
                        file_hash = None
                        try:
                            with open(pdf_path, 'rb') as f:
                                file_hash = hashlib.md5(f.read()).hexdigest()
                        except Exception as e:
                            logger.error(f"ファイルハッシュ計算中にエラーが発生しました: {pdf_path} - {str(e)}")
                        
                        # すでに処理済みのファイルはスキップ
                        if file_hash in processed_file_hashes:
                            logger.info(f"重複ファイルのためスキップします: {pdf_path}")
                            skipped_files += 1
                            continue
                        
                        logger.info(f"処理中: {pdf_path}")
                        
                        try:
                            # タイトルと著者を抽出
                            title, author = extract_title_and_author(pdf_path)
                            
                            # 新しいファイル名を作成
                            new_filename = f"{title}({author}).pdf"
                            # ファイル名のサニタイズ
                            new_filename = sanitize_filename(new_filename)
                            
                            # 出力先パス
                            output_path = os.path.join(output_folder, new_filename)
                            
                            # 同名ファイルが存在する場合は連番を付加
                            counter = 1
                            original_name = os.path.splitext(new_filename)[0]
                            while os.path.exists(output_path):
                                new_filename = f"{original_name}_{counter}.pdf"
                                output_path = os.path.join(output_folder, new_filename)
                                counter += 1
                            
                            # ファイルをコピー
                            shutil.copy2(pdf_path, output_path)
                            logger.info(f"コピー完了: {output_path}")
                            processed_files += 1
                            
                            # 処理済みハッシュに追加
                            if file_hash:
                                processed_file_hashes.add(file_hash)
                            
                            # 処理済みファイルを移動
                            processed_filename = os.path.basename(pdf_path)
                            processed_path = os.path.join(processed_folder, processed_filename)
                            
                            # 同名ファイルが移動先にある場合は連番を付加
                            counter = 1
                            while os.path.exists(processed_path):
                                base_name, ext = os.path.splitext(processed_filename)
                                processed_filename = f"{base_name}_{counter}{ext}"
                                processed_path = os.path.join(processed_folder, processed_filename)
                                counter += 1
                            
                            # ファイルを移動（コピー＋削除）
                            try:
                                shutil.move(pdf_path, processed_path)
                                logger.info(f"移動完了: {pdf_path} → {processed_path}")
                                moved_files += 1
                            except Exception as e:
                                logger.error(f"ファイル移動中にエラーが発生しました: {pdf_path} - {str(e)}")
                        
                        except Exception as e:
                            logger.error(f"ファイル処理中にエラーが発生しました: {pdf_path} - {str(e)}")
        
        logger.info(f"処理完了: 合計{total_files}ファイル中、{processed_files}ファイルを処理し、{moved_files}ファイルを移動しました。{skipped_files}ファイルはスキップされました。")
    
    except Exception as e:
        logger.error(f"実行中にエラーが発生しました: {str(e)}")

if __name__ == "__main__":
    main()
