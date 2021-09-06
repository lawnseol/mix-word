import pyexcel as pe
from itertools import product

def mix_array(keyword1s, keyword2s, keyword3s=None):
    items = []
    if (keyword3s == None):
        items = [keyword1s,keyword2s]
    else:
        items = [keyword1s,keyword2s,keyword3s]

    return list(product(*items))

def append(results, mixed_array):
    for row in mixed_array:
        keyword = (' '.join(str(e) for e in row))
        if (keyword.endswith(" ")):
            continue
        keyword = keyword.strip()
        if (keyword not in results):
            results.append(keyword)
    return results


if __name__ == "__main__":
    keyword = pe.get_array(file_name='keyword.xlsx')
    results = []

    results = append(results,mix_array(keyword[0],keyword[1]))
    results = append(results,mix_array(keyword[1],keyword[2]))
    results = append(results,mix_array(keyword[0],keyword[2]))
    results = append(results,mix_array(keyword[0],keyword[1],keyword[2]))

    print(results)

    pe.save_as(array=[results], dest_file_name="results.xlsx")




