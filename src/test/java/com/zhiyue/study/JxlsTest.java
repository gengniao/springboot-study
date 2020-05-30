package com.zhiyue.study;

import com.zhiyue.study.pojo.Employee;
import lombok.extern.java.Log;
import org.junit.Test;
import org.junit.runner.RunWith;
import org.jxls.common.Context;
import org.jxls.util.JxlsHelper;
import org.springframework.boot.test.context.SpringBootTest;
import org.springframework.test.context.junit4.SpringRunner;

import java.io.*;
import java.math.BigDecimal;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;

/**
 * @Description
 * @Project spirng-boot-jxls
 * @Author ZhiYue
 * @Date 2020/5/29 14:57
 */

@Log
@RunWith(SpringRunner.class)
@SpringBootTest(classes = SampleApplication.class)
public class JxlsTest {


    /**
     * jxls 基于excle模板导出excel
     * @throws IOException
     */
    @Test
    public void jxlsTest01() throws IOException {
        log.info("Running jxls excel generate demo");
        List<Employee> employees = getEmployees();

        String path = new File("").getAbsolutePath() + "/src/main/resources/aaa.xls";
        System.out.println("------------------------------");
        System.out.println(path);

        InputStream is = JxlsTest.class.getResourceAsStream("aaa.xls");
        //InputStream is = new FileInputStream(path);
        OutputStream os = new FileOutputStream("bbb.xls");
        Context context = new Context();
        context.putVar("employees", employees);

        JxlsHelper.getInstance().processTemplate(is, os, context);

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
