package com.zhiyue.study;

import com.zhiyue.study.pojo.Employee;
import jdk.nashorn.internal.runtime.logging.Logger;
import org.jxls.common.Context;
import org.jxls.util.JxlsHelper;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RestController;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import java.io.*;
import java.math.BigDecimal;
import java.net.URLEncoder;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;

/**
 * @Description jxls excel 下载
 * @Project spirng-boot-jxls
 * @Author ZhiYue
 * @Date 2020/5/29 17:39
 */
@RestController
public class JxlsController {

    @GetMapping("/export")
    public void jxlsExport(HttpServletRequest request, HttpServletResponse response) throws IOException {
        List<Employee> employees = getEmployees();
        String path = new File("").getAbsolutePath() + "/src/main/resources/aaa.xls";

        // InputStream is = JxlsTest.class.getResourceAsStream("aaa.xls");
        InputStream is = new FileInputStream(path);
        String fileName = URLEncoder.encode("export.xls", "UTF-8");
        response.setHeader("Content-Disposition", "attachment;filename=" + fileName);
        OutputStream os = response.getOutputStream();
        Context context = new Context();
        context.putVar("employees", employees);

        JxlsHelper.getInstance().processTemplate(is, os, context);



        os.flush();
        os.close();
        
    }


    private List<Employee> getEmployees() {
        List<Employee> data = new ArrayList<Employee>(2);
        Employee e1 = new Employee();
        e1.setBirthDate(new Date());
        e1.setBonus(new BigDecimal(200));
        e1.setName("zhiyue");
        e1.setPayment(new BigDecimal(5000));
        Employee e2 = new Employee();
        e2.setBirthDate(new Date());
        e2.setBonus(new BigDecimal(200));
        e2.setName("gengniao");
        e2.setPayment(new BigDecimal(5000));
        data.add(e1);
        data.add(e2);
        return data;
    }

}
