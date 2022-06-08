# C:\Users\jonson\AppData\Roaming\Anki2\test


from anki.collection import Collection
from anki.importing import TextImporter
import pandas as pd


Import_File = r'import.xlsx'
df = pd.read_excel(Import_File) #place "r" before the path string to address special character, such as '\'. Don't forget to put the file name at the end of the path + '.xlsx'

def add_new_card(col, deck_name, notetype_name, english_word, chinese_word):
    add_deck = col.decks.add_normal_deck_with_name(deck_name)
    print("add_deck:", add_deck)
    deck = col.decks.get(add_deck.id)

    notetype = col.models.by_name("綺英文") # 卡片
    deck["mid"] = notetype['id']
    col.decks.save(deck)

    a_new_note = col.new_note(notetype)
    a_new_note["英文單字"] = english_word
    a_new_note["中文"] = chinese_word

    col.add_note(a_new_note,add_deck.id)




col = Collection("collection.anki2")
# print(col.sched.deck_due_tree())



for ind in df.index:
    print(df['deck'][ind], df['english'][ind], df['chinese'][ind])
    add_new_card(col, df['deck'][ind], "綺英文", df['english'][ind], df['chinese'][ind])
    # add_new_card(col, "綺::level6::99表20", "綺英文", "english_word", "chinese_word")

col.close()




