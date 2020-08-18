from datetime import datetime
import os

today = str(datetime.now().date())
#1.Crawling Part
if not os.path.exists(f"./data/{today}"):
    os.makedirs(f"./data/{today}")
#1.1 Searching for images
#1.2 Saving images
#1.3 Duplication Prevention
#2. Generating PPT
#2.1 Loading Images to each Slide
#2.2 Image Fit
#2.3 Saving PPT as a video
#3. Youtube Upload
#3.1 Login to YOUTUBE
#3.2 Upload Generated video to YOUTUBE

#2.---––––––––––––––––––––––––––––––––––––––––
from pptx import Presentation
from pptx.util import Inches

prs = Presentation()
prs.slide_width = Inches(16)
prs.slide_height = Inches(9)
#디렉토리에서 사진 가져오기 
for item in os.listdir(f"./data/{today}"):
	img_path =os.getcwd()+'/data/'+today+"/"+item
	#print(img_path)
	blank_slide_layout = prs.slide_layouts[6] 
	slide = prs.slides.add_slide(blank_slide_layout)
	left = top = Inches(1)
	pic = slide.shapes.add_picture(img_path, Inches(0.5), Inches(1.75),
							   width=Inches(9), height=Inches(5))





#img_path = '/data/monty-truth.png'
filename = today+"_utub.pptx"

 #오늘 날짜를 가지고 파일이름 생성
prs.save(filename) #PPT저장 