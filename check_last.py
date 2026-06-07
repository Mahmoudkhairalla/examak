from docx import Document

doc = Document('/home/ubuntu/upload/Examak_IEEE_Conference_Paper.docx')
non_empty_paras = [p for p in doc.paragraphs if p.text.strip()]
print(f"Total non-empty paragraphs: {len(non_empty_paras)}")
for i in range(max(0, len(non_empty_paras)-5), len(non_empty_paras)):
    print(f"Para {i}: {non_empty_paras[i].text}")
