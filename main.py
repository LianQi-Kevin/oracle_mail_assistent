import base64
import xml.etree.ElementTree as ET
from dataclasses import dataclass
from datetime import datetime, timedelta
from typing import Literal
import re
import openpyxl
import threading

import requests

from config import config


ACCESS_TOKEN_LOCK = threading.Lock()

MAIN_RE = re.compile(
    r'^SLDS-BCEG-'  # 固定前缀
    r'(?P<unit>\d{3})-'  # 单体编号
    r'SDS-'
    r'(?P<discipline>[A-Z]+)-'  # 专业代码
    r'(?P<drawing>[A-Z0-9]+)'  # 图纸编号

    # ----------① 版本号（可选）----------
    r'(?:_*(?P<ver>[A-Z]|\d+\+[A-Z]|\d+[A-Z]|\d+))?'

    # ----------② 图名（可选）------------
    r'(?:\s+(?P<title>.+))?'

    r'\s*$'  # 行尾允许空白
)

VER_RE = re.compile(
    r'^(?:(?P<num>\d+)(?:\+(?P<plus_letter>[A-Z])|(?P<letter>[A-Z])?)'
    r'|(?P<pure_letter>[A-Z]))$'
)


@dataclass
class responseMailInfo:
    mailID: int
    MailNo: str
    SentDate: datetime
    subject: str
    AllAttachmentCount: int


@dataclass
class patternInfo:
    unit: str
    discipline: str
    drawing: str
    ver: str
    title: str


def clean_str(s: str) -> str:
    """清理字符串"""
    return re.sub(r'\s+', ' ', s).replace('＿', '_').strip()


def sortMailsByVer(mails: list[responseMailInfo]) -> list[responseMailInfo]:
    """
    Sort the mails by version
    """
    def ver_key(ver: str) -> tuple[int, int, int]:
        _m = VER_RE.match(ver) if ver else None
        if not _m:
            return 0, 4, 0  # 异常值，永远最后

        # ── 单字母 ──
        if _m.group('pure_letter'):  # L3
            return 0, 3, -ord(_m.group('pure_letter'))

        num_rank = -int(_m.group('num'))  # 数字降序
        if _m.group('plus_letter'):  # L0
            return num_rank, 0, -ord(_m.group('plus_letter'))
        if _m.group('letter'):  # L1
            return num_rank, 1, -ord(_m.group('letter'))
        return num_rank, 2, 0  # L2 纯数字

    def mail_ver_key(mail):
        _m = MAIN_RE.match(mail.subject)
        ver = _m.group('ver') if (_m and _m.group('ver')) else ''
        return ver_key(ver)

    return sorted(mails, key=mail_ver_key)


def responseClean(mail_response: list[responseMailInfo], search_params: patternInfo) -> list[responseMailInfo]:
    """
    Clean the mail response by filtering based title
    """
    subject = f"SLDS-BCEG-{search_params.unit}-SDS-{search_params.discipline}-{search_params.drawing}"
    return sortMailsByVer([mail for mail in mail_response if mail.subject.startswith(subject)])


def requestToken() -> dict:
    """
    Request an OAuth2 token using client credentials grant type

    https://help.aconex.com/zh/apis/implement-smart-construction-platform-oauth/#Implement-OAuth-in-a-User-Bound-Integration
    """

    def basicAuthEncode(client_id: str, client_secret: str) -> str:
        """
        Encode client_id and client_secret in base64 for Basic Auth

        https://en.wikipedia.org/wiki/Basic_access_authentication
        """
        auth_str = f"{client_id}:{client_secret}"
        b64_encoded = base64.b64encode(auth_str.encode()).decode()
        return b64_encoded

    response = requests.post(url=f"{config.lobby_url}/auth/token",
                             headers={"Content-Type": "application/x-www-form-urlencoded",
                                      "Authorization": f"Basic {basicAuthEncode(config.client_id, config.client_secret)}"},
                             data={"grant_type": "client_credentials", "user_id": config.aconex_user_id,
                                   "user_site": config.aconex_instance_url},
                             proxies={"http": "127.0.0.1:52538", "https": "127.0.0.1:52538"})

    response.raise_for_status()

    config.access_token = response.json().get("access_token")
    config.access_token_expires = datetime.now() + timedelta(seconds=response.json().get("expires_in") - 300)
    return response.json()


def searchMail(search_params: patternInfo = patternInfo(unit="000", discipline="A", drawing="A001", ver="_A", title=""),
               mail_box: Literal["INBOX", "SENTBOX", "ALL"] = "ALL") -> list[responseMailInfo]:
    """
    Search mails by subject

    https://help.aconex.com/zh/apis/mail-api-developer-guide/
    """

    def searchQueryCreator() -> str:
        """
        根据深化图编号规则生成 search_query 字符串
        unit     单体号, 3 位 (e.g. '001')
        discipline     专业代码, 大写 (e.g. 'HV')
        drawing  图号, 3 位 (e.g. '001')
        ver      版本号, 形如 '_0', '_A'；默认为 '*' 通配全部版本
        """
        # Lucene 表达式：拆分词后用 AND 组合 + 通配符
        tokens = f"{search_params.unit} {search_params.discipline} {search_params.drawing}{search_params.ver}".split()
        subject_cond = " AND ".join(f"{tok}" for tok in tokens)
        query = rf"subject:({subject_cond}) AND corrtypeid:23"
        # query = rf"subject:({subject_cond})"
        return query

    def responseMailInfoPostprocess(response_content) -> list[responseMailInfo]:
        """
        处理返回的邮件数据，转换为 responseMailInfo 列表
        """
        root = ET.fromstring(response_content.decode("utf-8"))
        export_list = []
        for _mail in root.find('SearchResults').iter('Mail'):
            # print(ET.tostring(m, encoding='utf-8').decode('utf-8'))
            export_list.append(responseMailInfo(mailID=int(_mail.attrib['MailId']), MailNo=_mail.findtext('MailNo'),
                                                SentDate=datetime.fromisoformat(_mail.findtext('SentDate').rstrip('Z')),
                                                subject=_mail.findtext('Subject'),
                                                AllAttachmentCount=int(_mail.findtext('AllAttachmentCount')), ))
        return export_list

    # 检查输入变量
    print(f"Search params: {search_params.__dict__}, mail box: {mail_box}")

    # 检查当前时间是否超过了config.access_token_expires，如果超过则调用requestToken()刷新token
    if not config.access_token_expires or not config.access_token or datetime.now() >= config.access_token_expires:
        with ACCESS_TOKEN_LOCK:
            print("Access token expired, refreshing...")
            requestToken()

    mail = []

    # SentBox
    if mail_box == "SENTBOX" or mail_box == "ALL":
        response = requests.get(
            url=f"{config.resource_url}/api/projects/{config.project_id}/mail",
            headers={"Authorization": f"Bearer {config.access_token}"},
            params={
                "mail_box": "SENTBOX",
                "search_query": searchQueryCreator(),
                "return_fields": "docno,subject,sentdate,allAttachmentCount,totalAttachmentsSize",
                "sort_field": "sentdate", "sort_direction": "DESC"
            },
            proxies={"http": "127.0.0.1:52538", "https": "127.0.0.1:52538"})

        response.raise_for_status()

        mail += responseMailInfoPostprocess(response.content)

    # InBox
    if mail_box == "INBOX" or mail_box == "ALL":
        response = requests.get(
            url=f"{config.resource_url}/api/projects/{config.project_id}/mail",
            headers={"Authorization": f"Bearer {config.access_token}"},
            params={
                "mail_box": "INBOX",
                "search_query": searchQueryCreator(),
                "return_fields": "docno,subject,sentdate,allAttachmentCount,totalAttachmentsSize",
                "sort_field": "sentdate", "sort_direction": "DESC"
            },
            proxies={"http": "127.0.0.1:52538", "https": "127.0.0.1:52538"})

        response.raise_for_status()
        mail += responseMailInfoPostprocess(response.content)

    # 使用mailID去重
    mail = list({_m.mailID: _m for _m in mail}.values())

    # mail 排序
    mail = responseClean(mail_response=mail, search_params=search_params)

    return mail


if __name__ == '__main__':
    # print(config)
    # requestToken()
    # print(searchMail(search_query="subject:(002 M M021) AND corrtypeid:23"))
    # _unit = "000"
    # _discipline = "A"
    # _drawing = "A002"
    #
    # print([mail for mail in searchMail(search_params=patternInfo(unit=_unit, discipline=_discipline, drawing=_drawing, ver="", title=""))])

    wb = openpyxl.load_workbook('test1.xlsx')
    for sheet in wb.worksheets:
        # sheet = wb['国泰瑞安']
        for row in sheet.iter_rows(min_row=2, max_col=9):
            if row[1].value is None or row[1] is None:
                continue
            m = MAIN_RE.match(clean_str(row[1].value))
            if not m:
                print("无法匹配:", row[1])
                continue

            matched_data: patternInfo.__dict__ = m.groupdict()
            cleaned_response = searchMail(search_params=patternInfo(
                unit=matched_data['unit'],
                discipline=matched_data['discipline'],
                drawing=matched_data['drawing'],
                ver="*", title=""
            ), mail_box="ALL")

            newest_mail = cleaned_response[0] if cleaned_response else None
            newest_matched_data = MAIN_RE.match(newest_mail.subject).groupdict() if newest_mail else None
            # matched_data['ver'] = MAIN_RE.match(newest_mail.subject).groupdict()['ver'] if newest_mail else ''
            # matched_data['title'] = MAIN_RE.match(newest_mail.subject).groupdict()['title'] if newest_mail else ''
            print(newest_matched_data)
            # ver
            row[4].value = newest_matched_data['ver'] if newest_matched_data else ''
            # sentDate
            row[5].value = newest_mail.SentDate.date().isoformat() if newest_mail else ''

    wb.save('test1_out.xlsx')
    wb.close()
