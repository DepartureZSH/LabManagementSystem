import cv2
import numpy as np
import re
from cnocr import CnOcr
from PIL import Image

class TableRecognition:
    def read_image(self, path):
        # 读取图片，以灰度模式
        img = cv2.imread(path, 1)
        # 显示图片，窗口名为image
        return img


    def to2(self, img):
        gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
        binary = cv2.adaptiveThreshold(~gray, 255, cv2.ADAPTIVE_THRESH_GAUSSIAN_C, cv2.THRESH_BINARY, 35, -5)
        return binary

    def Recognition(self, img_fp):
        img = self.read_image(img_fp)
        binary = self.to2(img)
        rows, cols = binary.shape
        scale = 40
        # 自适应获取核值
        # 识别横线:
        kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (cols // scale, 1))
        eroded = cv2.erode(binary, kernel, iterations=1)
        dilated_col = cv2.dilate(eroded, kernel, iterations=1)
        # 识别竖线：
        kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (1, rows // scale))
        eroded = cv2.erode(binary, kernel, iterations=4)
        dilated_row = cv2.dilate(eroded, kernel, iterations=1)
        # 将识别出来的横竖线合起来
        bitwise_and = cv2.bitwise_and(dilated_col, dilated_row)
        # 标识表格轮廓
        merge = cv2.add(dilated_col, dilated_row)
        # 两张图片进行减法运算，去掉表格框线
        merge2 = cv2.subtract(binary, merge)
        new_kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (2, 2))
        erode_image = cv2.morphologyEx(merge2, cv2.MORPH_OPEN, new_kernel)
        merge3 = cv2.add(erode_image, bitwise_and)
        # 将焦点标识取出来
        ys, xs = np.where(bitwise_and > 0)
        # 横纵坐标数组
        y_point_arr = []
        x_point_arr = []
        # 通过排序，排除掉相近的像素点，只取相近值的最后一点
        # 这个10就是两个像素点的距离，不是固定的，根据不同的图片会有调整，基本上为单元格表格的高度（y坐标跳变）和长度（x坐标跳变）
        i = 0
        sort_x_point = np.sort(xs)
        for i in range(len(sort_x_point) - 1):
            if sort_x_point[i + 1] - sort_x_point[i] > 10:
                x_point_arr.append(sort_x_point[i])
            i = i + 1
        # 要将最后一个点加入
        x_point_arr.append(sort_x_point[i])
        i = 0
        sort_y_point = np.sort(ys)
        for i in range(len(sort_y_point) - 1):
            if sort_y_point[i + 1] - sort_y_point[i] > 10:
                y_point_arr.append(sort_y_point[i])
            i = i + 1
        y_point_arr.append(sort_y_point[i])
        res = []
        # 循环y坐标，x坐标分割表格
        data = [[] for i in range(len(y_point_arr))]
        for i in range(len(y_point_arr) - 1):
            line = []
            for j in range(len(x_point_arr) - 1):
                # 在分割时，第一个参数为y坐标，第二个参数为x坐标
                cell = binary[y_point_arr[i]:y_point_arr[i + 1], x_point_arr[j]:x_point_arr[j + 1]]
                # 读取文字，此为默认英文
                ocr = CnOcr(cand_alphabet=["0", "1", "2", "3", "4", "5", "6", "7", "8", "9"])  # 所有参数都使用默认值
                out = ocr.ocr(cell)
                # print(out)
                # 去除特殊字符
                for each in out:
                    text = "".join(each[0])
                    line.append(text)
                j = j + 1
            res.append(line)
            i = i + 1
        return res
