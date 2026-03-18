# 图片解析修复方案

## 问题

原始代码使用了不正确的嵌套层级：
```python
for blip_fill in _iter_elements_by_tag(drawing, "blipFill"):
    for pic in _iter_elements_by_tag(blip_fill, "pic"):
        for pic_fill in _iter_elements_by_tag(pic, "picFills"):  # 这个标签不存在
            for blip in _iter_elements_by_tag(pic_fill, "blip"):
                # 处理图片
```

## 修复

正确的嵌套层级应该是：
```python
for blip_fill in _iter_elements_by_tag(drawing, "blipFill"):
    for blip in _iter_elements_by_tag(blip_fill, "blip"):  # 直接从 blipFill 查找 blip
        # 处理图片
```

## 需要修改的代码

将第139-141行：
```python
for pic in _iter_elements_by_tag(blip_fill, "pic"):
    for pic_fill in _iter_elements_by_tag(pic, "picFills"):
        for blip in _iter_elements_by_tag(pic_fill, "blip"):
```

修改为：
```python
for blip in _iter_elements_by_tag(blip_fill, "blip"):
```

这样就能正确解析图片了。