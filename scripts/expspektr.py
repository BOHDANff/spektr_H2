import matplotlib.pyplot as plt
experiment = [434, 443.2, 447.5, 462.8, 471, 486, 491.5, 501.2, 504.2, 579, 579.5, 581.5, 587.5, 588.5, 593.5, 603.4, 606.8, 609, 656, 656.5, 668]
ones = [1 for i in range(0, len(experiment))]
fig, ax = plt.subplots()
width_rectangle = 0.3
ax.bar(experiment, ones, width = width_rectangle)
ax.tick_params(axis='x', labelbottom='off', labeltop=False)
xlocs = [i for i in range(420, 690, 20)]

ax.set_xticks(xlocs, minor = False)


ax.set_ylabel('I, у.о.', fontsize=14)
ax.set_xlabel('$\lambda$, nm ' , fontsize=14)
plt.show()