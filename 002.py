
def frequencyOfMaxValue(numbers,q):
    n=len(numbers)
    ans = []
    for index in q:
        # Get the maximum in the subarray
        maxvalue=-1
        for i in range(index-1,n):
            if numbers[i]>maxvalue:
                maxvalue=numbers[i]
        # Get the count of maximum in the subarray
        count=0
        for i in range(index-1,n):
            if numbers[i]==maxvalue:
                count+=1
        ans.append(count)
    return ans


print(frequencyOfMaxValue([2, 2, 2], [1, 2, 3]))
