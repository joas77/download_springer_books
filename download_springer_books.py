import urllib.request
import urllib.parse
import urllib.error
import shutil
from openpyxl import load_workbook


def download_book(url, book_title, save_path):
    print(f"downloading {book_title}...")
    try:
        with urllib.request.urlopen(url) as stream:
            with open(f"{save_path}/{book_title}", "wb") as file:
                    shutil.copyfileobj(stream, file)
    except urllib.error.URLError as error:
        print(error)

if __name__ == "__main__":
    books = load_workbook("Free+English+textbooks.xlsx")
    ebooks_sheet = books.get_sheet_by_name('eBook list')

    filter = { "Computer Science",
        "Biomedical and Life Sciences",
        "Business and Economics",
        "Economics and Finance",
        "Engineering",
        "Intelligent Technologies and Robotics",
        "Mathematics and Statistics",
        "Physics and Astronomy"
    }

    for i in range(2, 408 + 1): #number of rows in excel
        book_title = ebooks_sheet["A" + str(i)].value
        book_title = book_title.replace("/", "-")
        book_title = book_title.strip()
        
        english_package_name = ebooks_sheet["L" + str(i)].value
        open_url = ebooks_sheet["S" + str(i)].value
        doi_url = ebooks_sheet["R" + str(i)].value
        _, doi = doi_url.split("http://doi.org/")

        pdf_link = "https://link.springer.com/content/pdf/"
        epub_link = "https://link.springer.com/download/epub/"

        doi_quote = urllib.parse.quote(doi, safe="")
        pdf_link += doi_quote + ".pdf"
        epub_link += doi_quote + ".epub"

        if english_package_name in filter:
            download_book(pdf_link, book_title+".pdf", "springer_books")
            download_book(epub_link, book_title+".epub", "springer_books")