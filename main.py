import base64
import xml.etree.ElementTree as ET
from dataclasses import dataclass
from datetime import datetime, timedelta
from typing import Literal

import requests

from config import config


@dataclass
class responseMailInfo:
    mailID: int
    MailNo: str
    SentDate: datetime
    subject: str
    AllAttachmentCount: int


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
    config.access_token_expires = datetime.now() + timedelta(seconds=response.json().get("expires_in") - 10)
    return response.json()


def searchMail(search_query: str = None, unit: Literal["000", "001", "002", "004"] = "000", prof: str = "A",
               drawing: str = "A001", ver: str = "*") -> list[responseMailInfo]:
    """
    Search mails by subject

    https://help.aconex.com/zh/apis/mail-api-developer-guide/
    """

    def searchQueryCreator() -> str:
        """
        根据深化图编号规则生成 search_query 字符串
        unit     单体号, 3 位 (e.g. '001')
        prof     专业代码, 大写 (e.g. 'HV')
        drawing  图号, 3 位 (e.g. '001')
        ver      版本号, 形如 '_0', '_A'；默认为 '*' 通配全部版本
        """
        # Lucene 表达式：拆分词后用 AND 组合 + 通配符
        tokens = f"{unit} {prof} {drawing}{ver}".replace("-", " ").replace("_", " ").split()
        subject_cond = " AND ".join(f"{tok}" for tok in tokens)
        query = rf"subject:({subject_cond}) AND corrtypeid:23"
        return query

    def responseMailInfoPostprocess(elem) -> responseMailInfo:
        """
        处理返回的邮件数据，转换为 responseMailInfo 列表
        """
        return responseMailInfo(mailID=int(elem.attrib['MailId']), MailNo=elem.findtext('MailNo'),
                                SentDate=datetime.fromisoformat(elem.findtext('SentDate').rstrip('Z')),
                                subject=elem.findtext('Subject'),
                                AllAttachmentCount=int(elem.findtext('AllAttachmentCount')), )

    # 检查当前时间是否超过了config.access_token_expires，如果超过则调用requestToken()刷新token
    if not config.access_token_expires or not config.access_token or datetime.now() >= config.access_token_expires:
        print("Access token expired, refreshing...")
        requestToken()

    # SentBox
    response = requests.get(
        url=f"{config.resource_url}/api/projects/{config.project_id}/mail",
        headers={"Authorization": f"Bearer {config.access_token}"},
        params={
            "mail_box": "sentbox",
            "search_query": searchQueryCreator() if search_query is None else search_query,
            "return_fields": "docno,subject,sentdate,allAttachmentCount,totalAttachmentsSize", "sort_field": "sentdate",
            "sort_direction": "DESC"
        },
        proxies={"http": "127.0.0.1:52538", "https": "127.0.0.1:52538"})

    response.raise_for_status()
    root = ET.fromstring(response.content.decode("utf-8"))
    mail = [responseMailInfoPostprocess(m) for m in root.find('SearchResults').iter('Mail')]

    # InBox
    response = requests.get(
        url=f"{config.resource_url}/api/projects/{config.project_id}/mail",
        headers={"Authorization": f"Bearer {config.access_token}"},
        params={
            "mail_box": "inbox",
            "search_query": searchQueryCreator() if search_query is None else search_query,
            "return_fields": "docno,subject,sentdate,allAttachmentCount,totalAttachmentsSize", "sort_field": "sentdate",
            "sort_direction": "DESC"
        },
        proxies={"http": "127.0.0.1:52538", "https": "127.0.0.1:52538"})

    response.raise_for_status()
    root = ET.fromstring(response.content.decode("utf-8"))
    mail += [responseMailInfoPostprocess(m) for m in root.find('SearchResults').iter('Mail')]

    # 使用mailID去重
    mail = list({m.mailID: m for m in mail}.values())

    return mail


if __name__ == '__main__':
    # print(config)
    # requestToken()
    # print(searchMail(unit="000", prof="ELV", drawing="ELV002", ver="_0"))
    print(searchMail(search_query="subject:(002 AND ELV002*) AND corrtypeid:23"))
