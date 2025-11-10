import matplotlib.pyplot as plt
from sklearn.datasets import make_blobs

# 1. 生成 4 簇，每簇标准差不一样
a, b = make_blobs(n_samples=800,
                  centers=[[0,0], [4,4], [-4,4], [0,-4]],
                  cluster_std=[0.8, 1.5, 0.5, 2.0],
                  random_state=42)

plt.scatter(a[:, 0], a[:, 1], c=b, cmap='viridis', s=25)
plt.title('make_blobs')
plt.show()