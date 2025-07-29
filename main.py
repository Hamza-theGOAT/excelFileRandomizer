import pandas as pd
import random


def randomizeLetters(s):
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
    df = pd.read_excel(
        'fileInput.xlsx',
        sheet_name='Sheet1'
    )
    print(f'Origina DataFrame:\n{df}')

    excluded = ['Name', 'Status']

    for col in df.columns:
        if col not in excluded:
            df[col] = df[col].astype(str).apply(randomizeLetters)

    print(f'Shuffled DataFrame:\n{df}')
    df.to_excel('fileOutput.xlsx', sheet_name='Sheet1', index=False)


if __name__ == '__main__':
    main()
