import pandas as pd
import numpy as np
import random
import os
from dotenv import load_dotenv


def randomizeLetters(s):
    # Set of letters and digits
    letters = [c for c in s if c.isalpha()]
    digits = [c for c in s if c.isdigit()]
    random.shuffle(letters)
    random.shuffle(digits)

    result = []
    for c in s:
        if c.isalpha():
            result.append(letters.pop())
        elif c.isdigit():
            result.append(digits.pop())
        else:
            result.append(c)

    return ''.join(result)


def main():
    # Get environment variables
    load_dotenv()
    fileInput = os.getenv('fileInput')
    shName = os.getenv('shName')
    excluded = os.getenv('excluded').split(',')

    # Get input file
    df = pd.read_excel(
        fileInput,
        sheet_name=shName
    )
    print(f'Origina DataFrame:\n{df}')
    print('_'*100)

    # Shuffle data within input file cols
    for col in df.columns:
        if col not in excluded:
            df[col] = df[col].replace(np.nan, '')
            df[col] = df[col].astype(str).apply(randomizeLetters)
    print(f'Shuffled DataFrame:\n{df}')

    # Save to excel
    df.to_excel('fileOutput.xlsx', sheet_name='Sheet1', index=False)


if __name__ == '__main__':
    main()
