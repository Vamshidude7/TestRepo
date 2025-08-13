def multiply(a, b):
    sign = -1 if (a < 0) ^ (b < 0) else 1
    a, b = abs(a), abs(b)
    if a.bit_count() < b.bit_count():
        a, b = b, a
    res = 0
    while b > 0:
        if b & 1:
            res += a
        a += a      # double a
        b >>= 1     # halve b
    return sign * res

a = 65454654654654998992222222222222228587878878787877
b = 11546465843121151261115454847687687687877867894897
c = multiply(a,b)
d= a*b
if c == d:
    print("Multiplication is correct")
else:
    print("Multiplication is incorrect")
print(multiply(a,b))  # Output: 12