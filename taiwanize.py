#import cn2tw
from pptx import Presentation
from langconv import Converter

def Traditional2Simplified(sentence): #繁體轉簡體
    sentence = str(sentence)
    sentence = Converter('zh-hans').convert(sentence)
    return sentence

def Simplified2Traditional(sentence): #簡體轉繁體
    sentence = str(sentence)
    sentence = Converter('zh-hant').convert(sentence)
    #sentence = Converter('zh-hant').convert(sentence)
    return sentence

def run( ): 
    prs = Presentation("example.pptx")
    for slide in prs.slides:            #檢視每個頁面
        for shape in slide.shapes:      #檢視每個框框
            if shape.has_text_frame:    #判斷該shape內是否有的文字
                text_frame = shape.text_frame  #獲取文字框
                TEST2 = Simplified2Traditional(text_frame.text) #文字繁體轉簡體
                print(TEST2)
                cur_text = text_frame.paragraphs[0].runs[0].text
                new_text = cur_text.replace(str(text_frame.text), str(TEST2)) #取代
                text_frame.paragraphs[0].runs[0].text = new_text #寫回
    prs.save("example2.pptx")
 
if __name__=="__main__":
    run( )