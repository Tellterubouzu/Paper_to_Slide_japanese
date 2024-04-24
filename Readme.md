# Paper to Slide Japanese
## How to use
### Openaiのapikyが必要です
.envファイルを作成し，そこにOPENAI_API_KEY = "your api key"
と記してください．一つの論文当たり50円前後かかります
### 環境構築
'''
python-m venv env
ev\Scripts\activate
pip install -r requirements.txt
'''
### path,発表者の名前の指定
./PDFフォルダに保存したpdfのパスをmain.py7行目のpath=に記してください
main.py92行目のsubtitle.text に発表者の名前を入れてください

### 実行
'''
python main.py
'''