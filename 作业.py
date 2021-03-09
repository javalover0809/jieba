import numpy as np
import pickle
import random
import matplotlib.pyplot as plt
from PIL import Image

path1 = '/Users/Oraida/Desktop/模式识别/img.pkl'

if __name__ == '__main__':
    with open(path1, 'rb') as fo:
        data = pickle.load(fo, encoding='bytes')
        data = np.array(data)
        # print(data)
        for i in range(0, 412):
            # print(data[i])
            for j in range(0, 549):
                # print(data[i][j])
                for k in range(0, 3):
                    data[i][j][k] = (data[i][j][k] - 100)/2
                    # print(data[i][j][k])
        print(data)
        print(data.shape)
        plt.imshow(data)
        plt.axis('off')
        plt.show()