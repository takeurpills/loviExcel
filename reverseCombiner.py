def reverseCombiner(rowList):
    # Don't do anything for empty list. Otherwise,
    # make a copy and sort.

    if len(rowList) == 0:
        return []
    sortedList = rowList[:]
    sortedList.sort()

    # Init, empty tuple, use first item for previous and
    # first in this run.

    tupleList = []
    firstItem = sortedList[0]
    prevItem = sortedList[0]

    # Process all other items in order.

    for item in sortedList[1:]:
        # If start of new run, add tuple and use new first-in-run.

        if item != prevItem + 1:
            tupleList = [(firstItem, prevItem + 1 - firstItem)] + tupleList
            firstItem = item

        # Regardless, current becomes previous for next loop.

        prevItem = item

    # Finish off the final run and return tuple list.

    tupleList = [(firstItem, prevItem + 1 - firstItem)] + tupleList
    return tupleList


# Test data, hit me with anything :-)

# myList = [1, 70, 71, 72, 98, 21, 22, 23, 24, 25, 99]

# Create tuple list, show original and that list, then process.

# tuples = reverseCombiner(myList)
# print(f"Original: {myList}")
# print(f"Tuples:   {tuples}\n")
# for tuple in tuples:
#     print(f"Would execute: worksheet.delete_rows({tuple[0]}, {tuple[1]})")
