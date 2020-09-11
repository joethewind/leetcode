import sys

class Solution(object):
    def twoSum(self, nums, target):
        length = len(nums)
        ans = []
        for i in range(0,length) :
            for j in range(i+1,length):
                if nums[i] + nums[j] == target:
                    return i,j
                else:
                    continue

if __name__ == '__main__': 
    a = Solution()
    b = [1,2,4]
    print(a.twoSum(b,3))

