import email
import io
import sys
import PyPDF2
from docx import Document
import xml.etree.ElementTree as ET

def parse_eml_file(eml_file_path):
    try:
        # Відкриття та розпарсування EML-файлу
        with open(eml_file_path, 'r', encoding='utf-8') as eml_file:
            msg = email.message_from_file(eml_file)

        # Отримання основних полів EML
        sender = msg.get('From')
        subject = msg.get('Subject')
        date = msg.get('Date')
        body = ""

        # Отримання тіла EML-повідомлення
        if msg.is_multipart():
            for part in msg.walk():
                content_type = part.get_content_type()
                content_disposition = str(part.get("Content-Disposition"))

                if "attachment" not in content_disposition:
                    payload = part.get_payload(decode=True)
                    if payload is not None:
                        body = payload.decode()

                    # Отримання текстового вмісту з HTML
                    if content_type == "text/html":
                        try:
                            import html2text
                            converter = html2text.HTML2Text()
                            converter.ignore_links = True
                            body = converter.handle(body)
                        except ImportError:
                            pass
        else:
            payload = msg.get_payload(decode=True)
            if payload is not None:
                body = payload.decode()

        # Отримання вкладень
        attachments = []
        for part in msg.walk():
            if part.get_content_maintype() == "multipart":
                continue
            filename = part.get_filename()
            if filename:
                content_type = part.get_content_type()
                payload = part.get_payload(decode=True)
                if payload is not None:
                    attachments.append((filename, content_type, payload))

        # Виведення результатів
        print(f"Відправник: {sender}")
        print(f"Тема: {subject}")
        print(f"Дата: {date}")
        print("Тіло повідомлення:")
        print(body)

        for filename, content_type, payload in attachments:
            print("\nВкладення:")
            print(f"Ім'я файлу: {filename}")
            print(f"Тип файлу: {content_type}")

            if content_type == "application/pdf":
                pdf_text = extract_text_from_pdf(payload)
                print("Вміст PDF:")
                print(pdf_text)
            elif content_type in ["application/msword", "application/vnd.openxmlformats-officedocument.wordprocessingml.document"]:
                doc_text = extract_text_from_doc(payload)
                print("Вміст DOC/DOCX:")
                print(doc_text)
            elif content_type == "text/xml":
                xml_text = extract_text_from_xml(payload)
                print("Вміст XML:")
                print(xml_text)
            else:
                print("Не підтримується розпарсування цього типу вкладення.")

    except Exception as e:
        print(f"Помилка: {str(e)}")

def extract_text_from_pdf(pdf_content):
    pdf_text = ""
    try:
        pdf_reader = PyPDF2.PdfReader(io.BytesIO(pdf_content))
        for page in pdf_reader.pages:
            pdf_text += page.extract_text()
    except Exception as e:
        pdf_text = f"Помилка при розпарсуванні PDF: {str(e)}"
    return pdf_text

def extract_text_from_doc(doc_content):
    doc_text = ""
    try:
        doc = Document(io.BytesIO(doc_content))
        for paragraph in doc.paragraphs:
            doc_text += paragraph.text + "\n"
    except Exception as e:
        doc_text = f"Помилка при розпарсуванні DOC/DOCX: {str(e)}"
    return doc_text

def extract_text_from_xml(xml_content):
    xml_text = ""
    try:
        root = ET.fromstring(xml_content)
        xml_text = ET.tostring(root, encoding='utf-8').decode()
    except Exception as e:
        xml_text = f"Помилка при розпарсуванні XML: {str(e)}"
    return xml_text
# Запуск

if __name__ == "__main__":
    if len(sys.argv) != 2:
        print("Використання: python parse_eml.py [шлях_до_файлу.eml]")
    else:
        eml_file_path = sys.argv[1]
        parse_eml_file(eml_file_path)
