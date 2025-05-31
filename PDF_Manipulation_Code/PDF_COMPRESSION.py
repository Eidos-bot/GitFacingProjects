from pdf2image import convert_from_path
import os

# The variable input_path_pdf is the file path of what is being compressed.
input_pdf_path = r""

size_red = 1
# Can only be an integer.
temp_folder = r"temp_images"
print("Converting...")
pdf_image = convert_from_path(input_pdf_path, use_pdftocairo=True, output_folder=temp_folder,
                              poppler_path=r"external_items/poppler-24.08.0/Library/bin")
print("Done converting, now compressing.")
new_pdf_image = []
itm_num = 1
for image in pdf_image:
    width, height = image.size
    new_size = (width//size_red, height//size_red)
    resized_image = image.resize(new_size)
    new_pdf_image.append(resized_image)
    # Very, very interestingly enough, duplicates are recognized in the list. So it will say,
    # the same index value multiple times through the log because it's just using the copy. This can be used somehow.
    # I'll change this, so it's not confusing but pdf_image.index(image) will let you know if there are duplicate pages.
    print(f"Page {itm_num} of {len(pdf_image)} compressed.")
    itm_num += 1

output_pdf_path = input_pdf_path.replace(".pdf", " compressed.pdf")

new_pdf_image[0].save(
    output_pdf_path, "PDF", resolution=100.0, save_all=True, append_images=new_pdf_image[1:]
)
for filename in os.listdir(temp_folder):
    file_path = os.path.join(temp_folder, filename)

    # Check if it's a file
    if os.path.isfile(file_path):
        # Delete files
        # This is because since you'd normally only compress for really large files, you'd likely get a lot of images in
        # the folder. This cleans it up after everything is done.
        os.remove(file_path)
original_file_size = os.path.getsize(input_pdf_path)
compressed_file_size = os.path.getsize(output_pdf_path)

change = "unchanged, staying"
if original_file_size > compressed_file_size:
    change = "reduced to"
elif compressed_file_size > original_file_size:
    change = "increased to"
percentage = f"{round(compressed_file_size / original_file_size, 2) * 100}%"

print(f"Original file size of {original_file_size} has been {change} {compressed_file_size}. ~{percentage} of the original.")
