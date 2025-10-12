"""
对于main函数测试新增函数功能

手动启动aria2c RPC服务端：
./aria2c.exe --enable-rpc --rpc-listen-all=false --rpc-listen-port=12768 --rpc-allow-origin-all --continue --save-session=./downloads/aria2.session --file-allocation=falloc
"""
import html
import re
import xml.etree.ElementTree as ET
from datetime import datetime, timezone, timedelta
from typing import Optional, Union

import openpyxl
from bs4 import BeautifulSoup
import aria2p

from pathlib import Path

from config import config
from dataclass import MailDetail, RegisteredDocumentAttachment, FromUserDetails, Recipient
from main import requestToken, clean_str, get_with_retry

XLSX_PATH = r"./图纸进度跟踪表_download.xlsx"

# aria2p
DOWNLOAD_PATH = Path("./downloads").resolve()   # 下载目录
RPC_PORT = 12768    # aria2c RPC 端口
RPC_SECRET = ""     # aria2c RPC 密钥（留空则不使用密钥）

ARIA2P_API = aria2p.API(aria2p.Client(host="http://localhost", port=RPC_PORT, secret=RPC_SECRET))


def viewMailMetadata(mail_id: Union[str, int]) -> MailDetail:
    """获取邮件元数据"""
    def _parse_datetime(dt: str) -> Optional[datetime]:
        """UTC ↔ +08:00 转换（保留毫秒）"""
        _TZ_CN = timezone(timedelta(hours=8))  # 东八区

        if not dt:
            return None
        utc_dt = datetime.fromisoformat(dt.replace("Z", "+00:00"))
        return utc_dt.astimezone(_TZ_CN)

    def _get_text(node, tag: str, default: str = "") -> str:
        """安全读取子节点文本，避免 .text 为 None 报错"""
        child = node.find(tag) if node is not None else None
        return child.text.strip() if child is not None and child.text else default

    def _html_to_text(raw: str) -> str:
        """
        使用 BeautifulSoup 把 MailData 里的富文本 HTML ➟ 纯文本。
        - <p>、<br> 等标签自动转换为换行
        - &lt; &gt; 实体自动解码
        """
        if not raw:
            return ""
        decoded = html.unescape(raw)  # 把 &lt; 之类实体转回 <
        soup = BeautifulSoup(decoded, "lxml")  # 速度更快；缺省回退 html.parser
        text = soup.get_text(strip=True)  # 保留换行
        text = re.sub(r"\s*\n\s*", "\n", text)  # 压缩相邻空行
        text = re.sub(r"[ \t]{2,}", " ", text)  # 连续空格→单空格
        return text.strip()

    def postprocess(xml_text: bytes) -> MailDetail:
        root = ET.fromstring(xml_text.decode("utf-8"))

        # ----- 附件列表 -----
        attachments = [RegisteredDocumentAttachment(attachment_id=a.attrib.get("attachmentId"),
                                                    document_no=_get_text(a, "DocumentNo"),
                                                    file_name=_get_text(a, "FileName"),
                                                    file_size=_get_text(a, "FileSize"),
                                                    title=_get_text(a, "Title"),
                                                    revision=_get_text(a, "Revision"),
                                                    document_id=_get_text(a, "DocumentId"),
                                                    ) for a in
                       (root.find("Attachments") or [])]

        # ----- 收件人列表 -----
        recipients = [Recipient(name=_get_text(r, "Name"), organization_name=_get_text(r, "OrganizationName"), ) for r
                      in (root.find("ToUsers") or [])]

        # ----- 发件人 -----
        fu = root.find("FromUserDetails")
        from_user_details = FromUserDetails(name=_get_text(fu, "Name"),
                                            organization_name=_get_text(fu, "OrganizationName"), )

        # ----- 组装 MailDetail -----
        return MailDetail(mail_id=root.attrib.get("MailId"), subject=_get_text(root, "Subject"),
                          sent_date=_parse_datetime(_get_text(root, "SentDate")),
                          mail_data=_html_to_text(_get_text(root, "MailData")), from_user_details=from_user_details,
                          attachments=attachments, recipients=recipients, )

    response = get_with_retry(url=f"{config.resource_url}/api/projects/{config.project_id}/mail/{mail_id}",
                              headers={"Authorization": f"Bearer {config.access_token}"})
    response.raise_for_status()
    return postprocess(xml_text=response.content)


def download_attachment_aria2c(attachment: RegisteredDocumentAttachment, subject: str, mail_id: str, sub_path: Optional[str] = None):
    """下载邮件附件"""
    global ARIA2P_API

    def build_options(sub_dir, file_name: str):
        target_path = DOWNLOAD_PATH / sub_dir if sub_path is None else DOWNLOAD_PATH / sub_path / sub_dir
        target_path.mkdir(parents=True, exist_ok=True)
        return {
            "dir": str(target_path),  # 下载目录
            "out": file_name,  # 保存的文件名
            "continue": "true",  # 断点续传
            "max-connection-per-server": "16",  # 每个服务器的最大连接数
            "split": 16,  # 文件分片数
            "min-split-size": "1M",  # 最小分片大小
            "file-allocation": "falloc",  # 文件预分配方式
            "header": [f"Authorization: Bearer {config.access_token}"],  # Aconex 附件需要
        }

    # 构造下载链接
    url = f"{config.resource_url}/api/projects/{config.project_id}/mail/{mail_id}/attachments/{attachment.attachment_id}"
    options = build_options(clean_str(subject), attachment.file_name)

    # 检查预期文件是否已经存在
    expected_file = Path(options['dir']) / options['out']
    if expected_file.exists():
        print(f"文件已存在，跳过下载: {expected_file}")
        return

    download = ARIA2P_API.add(url, options=options)[0]
    print(f"gid: {download.gid}, 保存到 {options['dir']}/{options['out']}")
    # return gid


if __name__ == '__main__':
    requestToken()

    wb = openpyxl.load_workbook(XLSX_PATH)
    sheet_list = ["建筑", "结构", "防水", "粗装", "1#楼精装", "泛光照明"]
    sheet = wb[sheet_list[5]]

    for row in sheet.iter_rows(min_row=2, max_col=50, values_only=True):
        if row[1] is not None and row[4].isdigit():
            data = {"id": row[1], "name": row[2], "ver": row[4], "mail_ID": row[8]}
            mail_response = viewMailMetadata(mail_id=data.get("mail_ID"))
            if "作废" in mail_response.subject:
                print("跳过作废邮件:", mail_response.subject)
                continue
            if "转发" in mail_response.subject:
                print("跳过转发邮件:", mail_response.subject)
                continue
            for att in mail_response.attachments:
                print(f"{mail_response.subject} 附件: {att.file_name} ({att.attachment_id})")
                download_attachment_aria2c(att, subject=mail_response.subject, mail_id=data.get('mail_ID'), sub_path=sheet.title)
