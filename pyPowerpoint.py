from pptx import Presentation
# テンプレート用ファイルの読み込み
prs = Presentation('/home/yuki/Downloads/trendTemplate.pptx')

name = "柏原 幸隆（かしはら ゆきたか）"
contents = \
"""
第２新卒エンジニア採用
趣味：プログラミング、読書
好きな製品：Deep Security
好きな技術：Python、Linux
マイブーム：Docker、Djnago

こちらのスライドはPythonで作成しました
Github：

"""

print("start")
for ns, slide in enumerate(prs.slides):
    for nsh, shape in enumerate(slide.shapes):
        if not shape.has_text_frame:
            continue
        for np, paragraph in enumerate(shape.text_frame.paragraphs):
            for rs, run in enumerate(paragraph.runs):
                if run.text  == 'title':
                    run.text = name
                
                if run.text == 'Contents':
                    run.text = contents

prs.save('/home/yuki/Downloads/test1.pptx')
print("end")

