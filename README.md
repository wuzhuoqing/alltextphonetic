# Add pinyin to all text using MS word.
The MS word has the Phonetic Guide dialog which can be used to add pinyin above chinese chars. However the default guide dialog only add pinyin to a few of the words instead of all text. We can instead use a macro to help with it. Here is [the macro code](https://github.com/wuzhuoqing/alltextphonetic/blob/master/alltext.bas). We simply need to create a new macro in WORD and paste the code into the macro editor, save, select all text and run the macro. For details please check the [Youtube video](https://youtu.be/RHq9e2dXghE) below.

The initial code is from [this stackoverflow](https://stackoverflow.com/questions/34602598/how-to-add-phonetic-guides-to-all-the-texts-at-once) and improved with batch for speed. The generated pinyin default alignment may not be what you want, You can use the field code search replace trick to change the formatting. (I learn the trick from another place but I lost the source link).

## Field code search replace trick for pinyin formatting
1. Select all text (easy shortcut is CTRL A).
2. Make the field codes visible with SHIFT F9.
3. Unselect the text.  [optional]
4. Do a find-and-replace operation (easy shortcut is CTRL H).  For example, search for `EQ \* jc2 \*` and replace with `EQ \* jc3 \*`. You can change the Phonetic Guide dialog setting and see what formatting field code it is generating for the style you want.
5. Select all text again (easy shortcut is CTRL A).
6. Make the field results visible either using SHIFT F9 (toggle field code visibility) or simply F9 (update field results).  Going to Print Preview might also work, depending upon your configuration.

# Youtube video link for the guide.
https://youtu.be/RHq9e2dXghE

# 给WORD中全部文字添加拼音
Word有个Phonetic Guide对话框可以用来给中文添加拼音。然而缺省的对话框Phonetic Guide只能对一小段文字添加拼音。我们可以通过使用宏代码来帮助解决这个问题。这里是用到的[宏代码 macro code](https://github.com/wuzhuoqing/alltextphonetic/blob/master/alltext.bas) 我们在WORD的宏编辑器中新建一个宏，把宏代码粘贴进去，保存，全选文字，运行宏即可。具体可以参照这个[Youtube 视频](https://youtu.be/RHq9e2dXghE)

原始代码来自[stackoverflow](https://stackoverflow.com/questions/34602598/how-to-add-phonetic-guides-to-all-the-texts-at-once) ，增加了批处理以提高速度。生成的拼音格式可以通过查找替换域代码的方式来调整(我是从别的地方看到这个技巧的，但忘了具体出处).

## 查找替换域代码的来调整拼音格式
1. 全选文字 (快捷键 CTRL A).
2. 按 SHIFT F9 显示域代码.
3. 取消选中文字.  [可选]
4. 查找替换 (快捷键 CTRL H).  例如, 查找 `EQ \* jc2 \*` 并替换为 `EQ \* jc3 \*`. 可以通过用Phonetic Guide对话框生成不同格式的拼音，查看对应的域代码来确定具体要替换成什么.
5. 再次全选文字 (快捷键 CTRL A).
6. 按 SHIFT F9 切换域代码.
