import sys
sys.path.insert(0, r"D:\PycharmProjects\hit-paper-helper")

from docx_manager.wps_ui.workflows.insert_image import insert_n_image_after_paragraph

DOCX = r"D:\PythonProjects\docx_manager\docx_manager\docx_engine\outputs\output.docx"

if __name__ == '__main__':
    items = [
        (r'D:\PythonProjects\docx_manager\image.png', 'tu1'),
        (r'D:\PythonProjects\docx_manager\image.png', 'tu2'),
        (r'D:\PythonProjects\docx_manager\image.png', 'tu3'),
        (r'D:\PythonProjects\docx_manager\image.png', 'tu4'),
    ]
    insert_n_image_after_paragraph(
        docx_path=DOCX,
        anchor_text='（1）气体静压轴承  加压气体经过节流器进入间隙，在间隙内产生压力气膜使物体浮起的气体轴承，结构如图1-1 (a)所示。……',
        items=items,
    )
    
