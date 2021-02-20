from tplreader import tpl_reader, tpl2excel, excel2word

def test_tpl_reader():
    word_path = "tpl.docx"
    config = tpl_reader(word_path)
    assert "项目简介" in config["items"]
    assert "Sheet1" in config["tables"]

def test_tpl2excel():
    word_path = "tpl.docx"
    tpl2excel(word_path)

def test_excel2word():
    excel_path, word_path, output_path = "confg1.xlsx", "tpl.docx", "output1.docx"
    excel2word(excel_path, word_path, output_path)