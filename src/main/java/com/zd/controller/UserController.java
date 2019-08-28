package com.zd.controller;

import com.zd.pojo.OrganizationImportInfo;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.beans.factory.annotation.Value;
import org.springframework.core.io.Resource;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RestController;

import javax.servlet.http.HttpServletResponse;
import java.io.IOException;
import java.io.OutputStream;
import java.util.ArrayList;
import java.util.List;

@RestController
@Slf4j
public class UserController {

    @Value("classpath:organization_import_template.xlsx")
    private Resource templateResource;

    @GetMapping("/download")
    public void downLoadExcel(HttpServletResponse response) throws Exception {
        if (templateResource == null) {
            log.error("IO异常，未读取到组织模板");
            throw new Exception("未读取到组织模板！");
        }

        List<OrganizationImportInfo> orgList = new ArrayList<>();
        buildOrgData(orgList);

        XSSFWorkbook fs = null;
        try {
            fs = new XSSFWorkbook(templateResource.getInputStream());
        } catch (IOException e) {
            log.error("IO异常，读取组织模板失败" + e);
            throw new Exception("读取组织模板失败！");
        }
        XSSFSheet sheet = fs.getSheetAt(0);
        int failSize = orgList.size();
        for (int i = 0; i < failSize; i++) {
            // 创建行(从第三行开始创建行，避开说明和标题行)
            XSSFRow row = sheet.createRow(i+2);
            // 一共13列
            for (int j = 0; j < 13; j++) {
                if (j == 0) {
                    row.createCell(j).setCellValue(orgList.get(i).getNotice());
                } else if (j == 1) {
                    row.createCell(j).setCellValue(orgList.get(i).getOrganizationCode());
                } else if (j == 2) {
                    row.createCell(j).setCellValue(orgList.get(i).getOrganizationName());
                } else if (j == 3) {
                    row.createCell(j).setCellValue(orgList.get(i).getAddress());
                } else if (j == 4) {
                    row.createCell(j).setCellValue(orgList.get(i).getCostCenterNumber());
                } else if (j == 5) {
                    row.createCell(j).setCellValue(orgList.get(i).getParentOrganizationCode());
                } else if (j == 6) {
                    row.createCell(j).setCellValue(orgList.get(i).getBrand());
                } else if (j == 7) {
                    row.createCell(j).setCellValue(orgList.get(i).getContactName());
                } else if (j == 8) {
                    row.createCell(j).setCellValue(orgList.get(i).getPhone());
                } else if (j == 9) {
                    row.createCell(j).setCellValue(orgList.get(i).getCustomNumber());
                } else if (j == 10) {
                    row.createCell(j).setCellValue(orgList.get(i).getCodCardNo());
                } else if (j == 11) {
                    row.createCell(j).setCellValue(orgList.get(i).getCustomerCode());
                } else {
                    row.createCell(j).setCellValue(orgList.get(i).getCheckCode());
                }
            }
        }
        response.setContentType("application/vnd.ms-excel");
        response.setHeader("Content-disposition", "attachment;filename=failMsg.xlsx");
        // 取得输出流
        OutputStream os = null;
        try {
            os = response.getOutputStream();
            fs.write(os);
        } catch (IOException e) {
            log.error("IO异常:{}", e);
            throw new Exception("导入出错!");
        }finally {
            if (os != null) {
                os.close();
            }
        }
    }

    private void buildOrgData(List<OrganizationImportInfo> orgList) {
        OrganizationImportInfo org1 = new OrganizationImportInfo();
        org1.setNotice("23123");
        org1.setOrganizationCode("123123");
        org1.setOrganizationName("adsf");
        org1.setAddress("asdf");
        org1.setCostCenterNumber("adsf");
        org1.setParentOrganizationCode("adsf");
        org1.setBrand("af");
        org1.setContactName("ff");
        org1.setPhone("f");
        org1.setCustomNumber("ss");
        org1.setCodCardNo("asdf");
        org1.setCustomerCode("asdf");
        org1.setCheckCode("asdf");

        OrganizationImportInfo org2 = new OrganizationImportInfo();
        org2.setNotice("1122");
        org2.setOrganizationCode("111");
        org2.setOrganizationName("233");
        org2.setAddress("3123");
        org2.setCostCenterNumber("123123");
        org2.setParentOrganizationCode("sss");
        org2.setBrand("asdf");
        org2.setContactName("asdf");
        org2.setPhone("asdf");
        org2.setCustomNumber("asdf");
        org2.setCodCardNo("ssss");
        org2.setCustomerCode("1111");
        org2.setCheckCode("223344");

        orgList.add(org1);
        orgList.add(org2);
    }

}
