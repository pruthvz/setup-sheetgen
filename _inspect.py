from docx import Document
from docx.oxml.ns import qn

def get_cell_text(tc):
    texts = []
    for elem in tc.iter():
        if elem.tag == qn('w:t') and elem.text:
            texts.append(elem.text)
    return ''.join(texts).strip()

def get_unique_cells(docpath):
    doc = Document(docpath)
    table = doc.tables[0]
    print('=== Unique cells in: ' + docpath + ' ===')
    for ri, row in enumerate(table.rows):
        tr = row._tr
        col = 0
        for tc in tr.findall(qn('w:tc')):
            text = get_cell_text(tc)
            tcPr = tc.find(qn('w:tcPr'))
            span = 1
            is_vmerge_cont = False
            if tcPr is not None:
                gs = tcPr.find(qn('w:gridSpan'))
                if gs is not None:
                    span = int(gs.get(qn('w:val'), '1'))
                vm = tcPr.find(qn('w:vMerge'))
                if vm is not None:
                    val = vm.get(qn('w:val'), '')
                    if val != 'restart':
                        is_vmerge_cont = True
            if not is_vmerge_cont:
                print('  [row=%d,col=%d,span=%d] %r' % (ri, col, span, text))
            col += span

get_unique_cells('_SU-Sheet-Laser.docx')
print()
get_unique_cells('_SU-Sheet-Combi-Aspex-S-F.docx')
