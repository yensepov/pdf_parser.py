def search_for_text(first_word, text):
    try:
        text_from_search = text
        start = text_from_search.find(first_word) + len(first_word) + 19
        end = start + 3
        result = text_from_search[start:end]
        return result
    except KeyError:
        pass

def search_for_text_in_block(doc, text):
    col = []
    for page in doc:
        blocks = page.get_text('blocks')
        for block in blocks:
            txt_from_block = block[4]
            if text in txt_from_block:
                col.append(txt_from_block)
    return col