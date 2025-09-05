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

#Below is the Java version of the same solution

# class Solution {
#    public int[] twoSum(int[] nums, int target) {
#        int n = nums.length;
#        Map<Integer,Integer> numMap = new HashMap<>();
#        for(int i = 0;i<n;i++){
#            int com = target - nums[i];
#            if(numMap.containsKey(com)){
#                return new int[]{numMap.get(com),i};
#            }
#            numMap.put(nums[i], i);
#       }
#        return new int[]{};
#    }
# }
