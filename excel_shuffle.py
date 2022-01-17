#!/usr/bin/env python
import numpy as np
import pandas as pd
import click


def read(file: str, skip=0):
    return pd.read_excel(file, skiprows=skip, header=None).astype(str)


@click.command()
@click.argument("infile")
@click.option(
    "-o",
    "--outfile",
    type=click.File(mode="w"),
    help="Optional file to write the results to. A '.xlsx' extension will produce a matrix-like "
    "file as INFILE, otherwise a plain text file ('.txt') is assumed.",
)
@click.option(
    "-n",
    "--max-results",
    default=None,
    type=int,
    help="Maximum number of results to write.",
)
@click.option(
    "-s", "--header-skip", default=2, help="Header rows to skip in INFILE. (default: 2)"
)
def main(infile: str, outfile: str, max_results: int, header_skip):
    df: pd.DataFrame = read(infile, skip=header_skip)
    n = max_results

    # mark invalids
    df[df.isin(["nan", ""])] = np.nan

    # find the longest row
    max_rows = max(df.count().max(), n or 0)

    # shuffle entries in columns
    shuffld = pd.DataFrame(index=range(max_rows))
    for c in df.columns:
        shuffld[c] = df[c].dropna().sample(max_rows, replace=True, ignore_index=True)

    # take n first entries and ensure type
    df = shuffld.head(n).astype(str)

    # make result strings
    s = pd.Series("", index=df.index)
    for c in df.columns:
        s += " " + df[c]

    # print out or write to file
    if outfile is None:
        for i in range(5 if n is None else n):
            print(s.iloc[i])
    elif outfile.endswith(".xlsx"):
        # we use df here !
        df.to_excel(outfile, index=False, header=False)
    else:
        if not outfile.endswith(".txt"):
            outfile += ".txt"
        with open(outfile, "w") as fh:
            for v in s:
                fh.write(v + "\n")


if __name__ == "__main__":
    main()
