def next_greater_element(arr):
    n = len(arr)
    result = [-1] * n
    stack = []
    for i in range(n-1,-1, -1):
        while stack and arr[stack[-1]] <= arr[i]:
            stack.pop()
        if stack:
            result[i] = arr[stack[-1]]
        
        stack.append(i)
    return result

arr = [2, 5, 3, 1, 2, 4, 6]
print(next_greater_element(arr))