import matplotlib.pyplot as plt
from matplotlib import cm
from mpl_toolkits.mplot3d import Axes3D
import numpy as np


def fun(pro, per):
	base = 1000
	pro_mid = 31163.04
	all_min = 500
	all_max = 2000
	beta = 0.01
	alpha = 0.05
	pro_delta = (pro_mid - pro) * alpha
	per_delta = (pro - per) * beta
	return np.clip(base+per_delta+pro_delta,all_min,all_max)

def Gini(A1,A2,A3,A4,A5):
	return 1 - (9*0.98*A1+7*0.93*A2+5*0.88*A3+3*0.83*A4+0.77*A5)/(5*(A1+A2+A3+A4+A5))

fun(31163.04,31163.04)


plt.rcParams['font.sans-serif']=['SimHei']
plt.rcParams['axes.unicode_minus']=False

fig1 = plt.figure(figsize=(16, 13))
ax = Axes3D(fig1)
X = np.arange(10000, 80000, 100)
Y = np.arange(8000, 80000, 100)
X, Y = np.meshgrid(X, Y)
Z = fun(X, Y)
plt.title("专项扣除额度")
ax.plot_surface(X, Y, Z, rstride=1, cstride=1, cmap=cm.coolwarm)
# ax.set_xlabel('地区收入中位数', color='r')
# ax.set_ylabel('个人收入', color='g')
# ax.set_zlabel('专项扣除额度', color='b')
ax.set_xlabel('地区收入中位数')
ax.set_ylabel('个人收入')
ax.set_zlabel('专项扣除额度')
plt.savefig('fig.png',bbox_inches='tight')
plt.show()
