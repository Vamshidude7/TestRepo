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

a = 7
b = 8
c = multiply(a,b)
d= a*b
if c == d:
    print("Multiplication is correct")
else:
    print("Multiplication is incorrect")
print(multiply(a,b))  # Output: 12