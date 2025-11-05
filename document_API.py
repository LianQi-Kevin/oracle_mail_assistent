from datetime import datetime, timezone, timedelta
from typing import Optional
from concurrent.futures import ThreadPoolExecutor, wait, ALL_COMPLETED

from requests import Response

from main import get_with_retry, requestToken
from config import config
import xml.etree.ElementTree as ET

from dataclasses import dataclass
import openpyxl


@dataclass
class DocumentInfo:
    title: str
    revision: str
    document_id: str
    document_number: str
    document_status: str    # "无效"则表示作废
    date_modified: datetime


@dataclass
class PageInfo:
    current_page: int
    page_size: int
    total_pages: int
    total_results: int
    total_results_on_page: int


def parseDatetime(dt: str) -> Optional[datetime]:
    """
    把 RFC-3339 / ISO-8601 字符串转为 datetime，并转换到 UTC+8。
    - 原始 API 字段形如 '2025-08-29T08:38:39.839Z'（Z 表示 UTC）
    - 返回值例如 2025-08-29 16:38:39.839+08:00
    """

    # 将 'Z' 替换为 '+00:00'，构造成可被 fromisoformat 解析的字符串
    utc_dt = datetime.fromisoformat(dt.replace("Z", "+00:00"))
    # 转换到东八区
    return utc_dt.astimezone(timezone(timedelta(hours=8)))


def list_registered_documents(search_query: str) -> list[DocumentInfo]:
    def _postprocess(xml_text: bytes) -> list[DocumentInfo]:
        _export_list: list[DocumentInfo] = []
        root = ET.fromstring(xml_text.decode("utf-8"))
        for _doc in root.find('SearchResults').iter('Document'):
            _export_list.append(DocumentInfo(
                title=_doc.findtext('Title'),
                revision=_doc.findtext('Revision'),
                document_id=_doc.attrib["DocumentId"],
                document_number=_doc.findtext('DocumentNumber'),
                document_status=_doc.findtext('DocumentStatus'),
                date_modified=parseDatetime(_doc.findtext('DateModified'))
            ))
        return _export_list

    def _get_page_info(xml_text: bytes) -> PageInfo:
        root = ET.fromstring(xml_text.decode("utf-8"))
        return PageInfo(
            current_page=int(root.attrib["CurrentPage"]),
            page_size=int(root.attrib["PageSize"]),
            total_pages=int(root.attrib["TotalPages"]),
            total_results=int(root.attrib["TotalResults"]),
            total_results_on_page=int(root.attrib["TotalResultsOnPage"])
        )

    def _get_response(page_size: int = 50, page_number: int = 1) -> Response:
        _response = get_with_retry(
            url=f"{config.resource_url}/api/projects/{config.project_id}/register",
            headers={"Authorization": f"Bearer {config.access_token}"},
            params={
                "search_query": search_query,
                "return_fields": "revision,discipline,docno,revisiondate,statusid,registered,title,doctype,reviewstatus,reviewSource",
                "sort_field": "revisiondate",
                "sort_direction": "DESC",
                "search_type": "PAGED",
                "page_size": page_size,
                "page_number": page_number
            })

        _response.raise_for_status()
        return _response

    response = _get_response()

    # support pagination
    page_info = _get_page_info(response.content)
    if page_info.total_pages > 1:
        all_docs: list[DocumentInfo] = _postprocess(response.content)

        # 构造线程池
        pool = ThreadPoolExecutor(max_workers=min(32, (os.cpu_count() or 1) * 5))
        all_tasks = []

        for page_num in range(2, page_info.total_pages + 1):
            response = _get_response(page_number=page_num)
            all_docs.extend(_postprocess(response.content))

        # 等待所有任务完成
        wait(all_tasks, return_when=ALL_COMPLETED)
        pool.shutdown()

        return all_docs

    return _postprocess(response.content)


if __name__ == '__main__':
    requestToken()
    registered_doc_list = list_registered_documents(search_query="SDS")
    print(len(registered_doc_list))
    # for doc in registered_doc_list:
    #     print(f"{doc.document_id} | {doc.document_number} | {doc.title} | {doc.revision} | {doc.document_status} | {doc.date_modified.isoformat()}")

    # wb = openpyxl.Workbook()
    # ws = wb.active
    # ws.title = "C2-4"
    # ws.append(["idx", "Title", "Revision", "Document Number"])
    # for idx, doc_info in enumerate(registered_doc_list):
    #     ws.append([idx + 1, doc_info.title, doc_info.revision, doc_info.document_number])
    #
    # wb.save("registered_C2-4_documents.xlsx")
    # wb.close()
