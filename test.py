import sys
sys.path.insert(0, r"D:\PycharmProjects\hit-paper-helper")

from docx_manager.wps_ui.workflows.insert_image import insert_n_images_one_col
from docx_manager.wps_ui.workflows.insert_two_images import insert_n_images_two_col

DOCX         = r"D:\PycharmProjects\hit-paper-helper\docx_manager\docx_engine\outputs\output.docx"
ANCHOR_IMAGE = r"D:\PycharmProjects\hit-paper-helper\anchor.png"
IMAGE        = r"D:\PycharmProjects\hit-paper-helper\image.png"

if __name__ == '__main__':
    mode = 1
    if mode == 1:
        insert_n_images_one_col(
            docx_path=DOCX,
            anchor_text='（1）气体静压轴承  加压气体经过节流器进入间隙，在间隙内产生压力气膜使物体浮起的气体轴承，结构如图1-1 (a)所示。……',
            anchor_image=ANCHOR_IMAGE,
            images=[IMAGE, IMAGE],
            captions=['子图题a', '子图题b'],
        )
    else:
        insert_n_images_two_col(
            docx_path=DOCX,
            anchor_text='（1）气体静压轴承  加压气体经过节流器进入间隙，在间隙内产生压力气膜使物体浮起的气体轴承，结构如图1-1 (a)所示。……',
            anchor_image=ANCHOR_IMAGE,
            images=[IMAGE, IMAGE, IMAGE, IMAGE],
            captions=['子图题a', '子图题b', '子图题c', '子图题d'],
            total_caption='图1-1 总图题',
            debug=True,           # 开启截图调试，截图存 debug/<timestamp>/
            run_phases=(1, 2, 3, 4, 5),  # 可改为单阶段如 (1,) 或 (2,) 单独测试
        )
        
