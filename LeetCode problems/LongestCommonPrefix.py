# Write a function to find the longest common prefix string amongst an array of strings.

# If there is no common prefix, return an empty string "".

 

# Example 1:

# Input: strs = ["flower","flow","flight"]
# Output: "fl"

# Example 2:

# Input: strs = ["dog","racecar","car"]
# Output: ""
# Explanation: There is no common prefix among the input strings.
 

# Constraints:

# 1 <= strs.length <= 200
# 0 <= strs[i].length <= 200
# strs[i] consists of only lowercase English letters if it is non-empty.
import collections
from abc import ABC, abstractmethod
class solution:
    @abstractmethod
    def longestCommonPrefix(self, strs):
        if not strs:
            return ""
        
        prefix = strs[0]
        
        for s in strs[1:]:
            while not s.startswith(prefix):
                prefix = prefix[:-1]
                if not prefix:
                    return ""
        
        return prefix

solution = solution()
print(solution.longestCommonPrefix(["flower","flow","flight"])+" "+"Vamshi")  # Output: "fl


# Below is the Java version of the same solution

# class Solution {
#     public String longestCommonPrefix(String[] strs) {
#         if (strs == null || strs.length == 0) {
#             return "";
#         }
#         String prefix = strs[0];
#         for (int i = 1; i < strs.length; i++) {
#             while (!strs[i].startsWith(prefix)) {
#                 prefix = prefix.substring(0, prefix.length() - 1);
#                 if (prefix.isEmpty()) {
#                     return "";
#                 }
#             }
#         }
#         return prefix;
#     }
# }
