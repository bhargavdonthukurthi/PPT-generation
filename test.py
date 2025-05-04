def distribute_paragraphs_to_slides(paragraph_lengths, lines_per_slide=25):
    paragraph_indices = sorted(range(len(paragraph_lengths)), key=lambda i: paragraph_lengths[i], reverse=True)
    slides = []
    remaining_lines = []

    for index in paragraph_indices:
        length = paragraph_lengths[index]
        placed = False
        for i in range(len(slides)):
            if remaining_lines[i] >= length:
                slides[i].append(index)
                remaining_lines[i] -= length
                placed = True
                break
        if not placed:
            slides.append([index])
            remaining_lines.append(lines_per_slide - length)

    return slides

# Example usage with your provided data:
paragraph_lengths = [4, 24, 18, 6, 5,17, 19,2 ,6,3,9,7,3,2,19]
slide_distribution = distribute_paragraphs_to_slides(paragraph_lengths)

for i, slide in enumerate(slide_distribution):
    print(f"Slide {i+1}: Paragraphs with original indices {slide}")

