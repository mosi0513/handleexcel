

while True:
		
	text = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"

	str = input("请输入字母查询下标数：")

	result = text.find(str) + 1

	print(result)
