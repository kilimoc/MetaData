theList = [[1,2,3], [4,5,6], [7,8,9]]
for i in range(len(theList)):
    if 5 in theList(i):
        print("[{0}][{1}]".format(i, theList[i].index(5))) #[1][1]