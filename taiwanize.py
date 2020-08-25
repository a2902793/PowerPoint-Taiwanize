from pptx import Presentation
from langconv import *

def Traditional2Simplified(sentence): #繁體轉簡體
    sentence = str(sentence)
    sentence = Converter('zh-hans').convert(sentence)
    return sentence

def Simplified2Traditional(sentence): #簡體轉繁體
    sentence = str(sentence)
    sentence = Converter('zh-hant').convert(sentence)
    return sentence

def run( ): 
    ppt = Presentation("example.pptx")
    for slide in ppt.slides:    #檢視每個頁面
	    for shape in slide.shapes:    #檢視每個文字框
		    if shape.has_text_frame:  #判斷該shape內是否有的文字
			    text_frame = shape.text   #獲取文字框
                TEST2 = Simplified2Traditional(TEST1)    #簡體轉繁體繁體還是繁體




			    print(TEST2) 

                

                

    print("幹")

    
if __name__=="__main__":
    run( )

                




                

                
            

    
