from main import get_with_retry, requestToken
from config import config
import xml.etree.ElementTree as ET

from dataclasses import dataclass
import openpyxl


@dataclass
class DocumentInfo:
    title: str
    revision: str
    document_number: str


def list_registered_documents(search_query: str) -> list[DocumentInfo]:
    def postprocess(xml_text: bytes) -> list[DocumentInfo]:
        _export_list: list[DocumentInfo] = []
        root = ET.fromstring(xml_text.decode("utf-8"))
        for _doc in root.find('SearchResults').iter('Document'):
            if "C2-4" in _doc.findtext('Title'):
                _export_list.append(DocumentInfo(
                    title=_doc.findtext('Title'),
                    revision=_doc.findtext('Revision'),
                    document_number=_doc.findtext('DocumentNumber')
                ))
        return _export_list

    response = get_with_retry(
        url=f"{config.resource_url}/api/projects/{config.project_id}/register",
        headers={"Authorization": f"Bearer {config.access_token}"},
        params={
            "search_query": search_query,
            "return_fields": "revision,discipline,docno,revisiondate,statusid,registered,title,doctype,reviewstatus,reviewSource",
            "sort_field": "revisiondate",
            "sort_direction": "DESC"
        })

    response.raise_for_status()
    return postprocess(response.content)


if __name__ == '__main__':
    requestToken()
    registered_doc_list = list_registered_documents(search_query="C2")

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "C2-4"
    ws.append(["idx", "Title", "Revision", "Document Number"])
    for idx, doc_info in enumerate(registered_doc_list):
        ws.append([idx + 1, doc_info.title, doc_info.revision, doc_info.document_number])

    wb.save("registered_C2-4_documents.xlsx")
    wb.close()



