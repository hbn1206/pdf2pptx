from PIL import Image

def convert_pdf_to_pptx(pdf_data):
    pdf_document = fitz.open("pdf", pdf_data)
    presentation = Presentation()

    for page_num in tqdm(range(len(pdf_document)), desc="Converting PDF to PPTX"):
        page = pdf_document.load_page(page_num)
        
        # 페이지 크기 가져오기
        page_width = Inches(page.rect.width / 72)
        page_height = Inches(page.rect.height / 72)
        
        # 슬라이드 크기 설정
        presentation.slide_width = page_width
        presentation.slide_height = page_height
        
        # PDF 페이지를 이미지로 렌더링
        pix = page.get_pixmap()
        image = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
        
        # 슬라이드에 이미지 추가
        slide = presentation.slides.add_slide(presentation.slide_layouts[5])
        
        image_width, image_height = image.size
        aspect_ratio = image_width / image_height

        if page_width / page_height > aspect_ratio:
            new_height = page_height
            new_width = new_height * aspect_ratio
        else:
            new_width = page_width
            new_height = new_width / aspect_ratio

        left = (page_width - new_width) / 2
        top = (page_height - new_height) / 2
        
        # 이미지를 슬라이드에 배치
        image_bytes = BytesIO()
        image.save(image_bytes, format="PNG")
        image_bytes.seek(0)
        slide.shapes.add_picture(image_bytes, left, top, width=new_width, height=new_height)
    
    # 프레젠테이션을 BytesIO에 저장하여 반환
    pptx_data = BytesIO()
    presentation.save(pptx_data)
    pptx_data.seek(0)
    return pptx_data
