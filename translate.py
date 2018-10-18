from HandleJs import Py4Js

def translate(content):
    if len(content) > 4891:
        print('翻译的长度超过4891 word 限制!\nPlease submit a shorter string')

    # 按照谷歌翻译页JavaScript代码生成验证tk验证码
    js =Py4Js()
    tk = js.getTk(content)

    param = {'tk': tk, 'q': content}

    result = requests.get("""http://translate.google.cn/translate_a/single?client=t&sl=en
        &tl=zh-CN&hl=zh-CN&dt=at&dt=bd&dt=ex&dt=ld&dt=md&dt=qca&dt=rw&dt=rm&dt=ss
        &dt=t&ie=UTF-8&oe=UTF-8&clearbtn=1&otf=1&pc=1&srcrom=0&ssel=0&tsel=0&kc=2""", params=param)
#    return(result.json())
    result = result.json()
    clear_result = ''
    for text in result[0][0:-1]:
        clear_result = clear_result + text[0]
    return(clear_result)
