import pandas as pd


papers = pd.read_csv(r"C:\Users\wxs\Desktop\SCI\savedrecs(2).txt",sep="\t",index_col=False)

print(papers.info())

print(papers['TI'][0])