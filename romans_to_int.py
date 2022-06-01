class Solution:

    def romanToInt(self, s: str) -> int:
        nums = {"IV":" 4 ", "IX":" 9 ", "XL":" 40 ", "XC":" 90 ", "CD":" 400 ", "CM":" 900 ", "I":" 1 ", "V":" 5 ", "X":" 10 ", "L":" 50 ", "C":" 100 ", "D":" 500 ", "M":" 1000 "}
        for key, value in nums.items():
            s = s.replace(key, value)
        res = 0
        for i in s.split():
            res += int(i)
        return res

    def intToRoman(self, num: int) -> str:
        nums = {1000:"M", 900:"CM", 500:"D", 400:"CD", 100:"C", 90:"XC", 50:"L", 40:"XL", 10:"X", 9:"IX", 5:"V", 4:"IV", 1:"I"}
        res = ""
        for i, char in nums.items():
            while num >= i:
                res += char
                num -=i
        return res




a = Solution()
print(a.romanToInt('MCMXCIV'))
print(a.intToRoman(1994))