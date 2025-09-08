import openpyxl
from main import MAIN_RE, searchMail, requestToken, searchWorkflow, clean_str
from dataclass import patternInfo

if __name__ == '__main__':
    wb = openpyxl.load_workbook('筑安机电-图纸清单.xlsx')
    sheet = wb.active

    requestToken()

    for row in sheet.iter_rows(min_row=2, max_row=74, max_col=50):
        # 读图纸编号
        print(row[0].value)
        matched = MAIN_RE.match(clean_str(row[0].value)).groupdict()
        print(matched)

        # 写出直接解析结果
        # row[16].value = rf"SLDS-BCEG-{matched['unit']}-SDS-{matched['discipline']}-{matched['drawing']}"    # 图纸编号
        # 查邮件
        mailResponse = searchMail(patternInfo(
            unit=matched['unit'],
            discipline=matched['discipline'],
            drawing=matched['drawing'],
        ), "ALL")

        if mailResponse:
            matched_new = MAIN_RE.match(clean_str(mailResponse[0].subject)).groupdict()
            print(matched_new)
            row[2].value = rf"SLDS-BCEG-{matched_new['unit']}-SDS-{matched_new['discipline']}-{matched_new['drawing']}"    # 图纸编号
            row[3].value = matched_new['title']    # 图纸名称
            row[4].value = mailResponse[0].SentDate.strftime('%Y-%m-%d')   # 发图时间
            row[5].value = matched_new['ver']  # 版本号
            row[6].value = matched_new.get("wf", "")  # 工作流编号

            # wf
            base_col = 8
            if matched_new["wf"]:
                wfResponse = searchWorkflow(matched_new["wf"])
                for workflow in wfResponse.workflows:
                    print(
                        f"Workflow ID: {workflow.workflow_id}, Step Status: {workflow.step_status}, Step Name: {workflow.step_name}, "
                        f"Step Out Come: {workflow.step_outcome}, Assignee: {workflow.assignees[0].name} "
                        # f"Organization Name: {workflow.assignees[0].organization_name}"
                    )
                    if workflow.step_name == "最终" and workflow.step_outcome != "正等待处理":
                        # 审核完成，写入最终审核状态
                        row[7].value = f"code {workflow.step_outcome.split('-')[0]}" if workflow.step_status != "已终止" else "工作流已终止"   # 审核状态
                    if workflow.step_outcome == "正等待处理":
                        # 正在处理的工作流，写入处理人和状态
                        row[base_col].value = workflow.assignees[0].organization_name.split(" ")[0]
                        row[base_col + 1].value = workflow.assignees[0].name.split(" ")[-1]
                        row[base_col + 2].value = workflow.step_status
                        base_col += 3

    wb.save('sample_export.xlsx')
    wb.close()
