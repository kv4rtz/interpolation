def to_power_symbol(x):
    superscript_digits = {
        0: '⁰', 1: '¹', 2: '²', 3: '³', 4: '⁴', 5: '⁵', 
        6: '⁶', 7: '⁷', 8: '⁸', 9: '⁹'
    }
    return ''.join(superscript_digits[int(digit)] for digit in str(x))
