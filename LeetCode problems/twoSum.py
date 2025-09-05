class Solution:
    def twoSum(self, nums, target):
        num_map = {}  # dictionary to store number -> index
        for i, num in enumerate(nums):
            com = target - num
            if com in num_map:
                return [num_map[com], i]
            num_map[num] = i
        return []
    
# Example usage:
solution = Solution()
print(solution.twoSum([2, 7, 11, 15], 9))
