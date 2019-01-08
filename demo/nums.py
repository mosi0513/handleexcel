
import random


k = 0 
for x in range(0,8):

	num = random.randint(1,2)
	num3 = random.randint(1,9)
	num2 = str(0.1) + str(num) + str(num3)

	content = float(num2)
	print(content)
	k +=content


print(k)