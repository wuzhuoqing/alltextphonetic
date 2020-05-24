# Add pinyin to all text using MS word.
The MS word has the Phonetic Guide dialog which can be used to add pinyin above chinese chars. However the default guide dialog only add pinyin to a few of the words instead of all text. We can instead use a macro to help with it. Here is [the macro code](https://github.com/wuzhuoqing/alltextphonetic/blob/master/alltext.bas)

The initial code is from [this stackoverflow](https://stackoverflow.com/questions/34602598/how-to-add-phonetic-guides-to-all-the-texts-at-once) and improved with batch for speed. The generated pinyin default alignment may not be what you want, You can use the field code search replace trick to change the formatting. (I learn the trick from another place but I lost the source link).

## Field code search replace trick for pinyin formatting
1. Select all text (easy shortcut is CTRL A).
2. Make the field codes visible with SHIFT F9.
3. Unselect the text.  [optional]
4. Do a find-and-replace operation (easy shortcut is CTRL H).  For example, search for `EQ \* jc2 \*` and replace with `EQ \* jc3 \*`. You can change the Phonetic Guide dialog setting and see what formatting field code it is generating for the style you want.
5. Select all text (easy shortcut is CTRL A).
6. Make the field results visible either using SHIFT F9 (toggle field code visibility) or simply F9 (update field results).  Going to Print Preview might also work, depending upon your configuration.

# Youtube video link for the guide.
https://youtu.be/RHq9e2dXghE
