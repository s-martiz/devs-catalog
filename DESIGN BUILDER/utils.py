import re, unicodedata

def tidy_columns(df, keep_accents=True, lowercase=True):
    """
    Nettoie les noms de colonnes d'un DataFrame :
      - corrige les encodages moisis (faÃ§ade -> façade) si possible
      - supprime espaces/ponctuation superflus, normalise en snake_case
      - garde (ou retire) les accents selon keep_accents
      - garantit l'unicité des noms (suffixes _2, _3, ...)
    """
    try:
        import ftfy
        fix_text = ftfy.fix_text
    except Exception:
        # fallback : tentative de réparation UTF8 mal décodé en latin-1
        def fix_text(s):
            try:
                return str(s).encode("latin1").decode("utf-8")
            except Exception:
                return str(s)

    def clean_one(col: str) -> str:
        s = fix_text(str(col))
        # espaces et caractères invisibles
        s = s.replace("\u00A0", " ").strip()
        s = re.sub(r"\s+", " ", s)
        # ponctuation légère -> underscore
        s = (s.replace("’", "'")
               .replace("'", "_")
               .replace("/", "_")
               .replace("-", "_"))
        # enlever () [] et autres
        s = re.sub(r"[()\[\]]", "", s)
        # espaces -> underscore
        s = s.replace(" ", "_")
        # accents ?
        if not keep_accents:
            s = unicodedata.normalize("NFKD", s).encode("ascii", "ignore").decode("ascii")
        # garder uniquement letters/digits/underscore
        s = re.sub(r"[^0-9A-Za-z_\u00C0-\u017F]", "", s)  # autorise accents si keep_accents=True
        # compacter les underscores
        s = re.sub(r"_+", "_", s).strip("_")
        if lowercase:
            s = s.lower()
        return s

    new_cols = [clean_one(c) for c in df.columns]

    # unicité
    seen = {}
    unique = []
    for c in new_cols:
        if c not in seen:
            seen[c] = 1
            unique.append(c)
        else:
            seen[c] += 1
            unique.append(f"{c}_{seen[c]}")
    df = df.copy()
    df.columns = unique
    return df
