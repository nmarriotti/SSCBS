def BinarySearch(array, x, start, end):
    ''' Locates a value stored in an ordered list '''
    if start > end:
        return False

    # Find middle element
    middle = int(start + ((end - start) /2))

    # Check if middle element contains x
    if array[middle] == x:
        return True
          
    # Slice list based on where x falls in the sorted array (left or right)
    elif x.lower() < array[middle].lower():
        return BinarySearch(array, x, start=start, end=middle-1)
    else:
        return BinarySearch(array, x, start=middle+1, end=end)