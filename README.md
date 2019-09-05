Simple script to find exact locations of non-ASCII characters in Excel spreadsheets, including cell coordinates and character positions in strings.
Just a few lines.

Example usage:

    $ python find_nonascii_xls.py form.xlsx 
    Finding non-ASCII cells and characters ...

    <Cell 'survey'.E78>
        not(selected(., 'None/Don’ttransport') and count-selected(.) >=2)
        character: ’
        index pos: 25

    <Cell 'choices'.C89>
        Agrovet–Agri Outputs
        character: –
        index pos: 7

    <Cell 'choices'.B183>
        None/Don’ttransport
        character: ’
        index pos: 8

    <Cell 'choices'.C183>
        None/Don’t transport
        character: ’
        index pos: 8

QED | https://qed.ai
