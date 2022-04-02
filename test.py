
a = input("請輸入權數:")
b = input("請輸入號碼:")
# a = "123456"
# b = "987654"
h = []
j = []
s = 0

print("開始了")

# for i in a:
#     c = int(float(i))
#     h.append(c)

# for i in b:
#     c = int(float(i))
#     j.append(c)

for i in range(len(a)):
    g = int(float(a[i]))
    d = int(float(b[i]))
    h.append(g)
    j.append(d)

    x = h[i]*j[i]
    s = s+x
# print(s)
# v = s % 10
# print(v)
result = 10-(s % 10)
# result = 10-v
print(result)
