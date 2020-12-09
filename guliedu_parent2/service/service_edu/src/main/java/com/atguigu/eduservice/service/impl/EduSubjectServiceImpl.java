package com.atguigu.eduservice.service.impl;

import com.atguigu.eduservice.entity.EduSubject;
import com.atguigu.eduservice.mapper.EduSubjectMapper;
import com.atguigu.eduservice.service.EduSubjectService;
import com.baomidou.mybatisplus.core.conditions.query.QueryWrapper;
import com.baomidou.mybatisplus.extension.service.impl.ServiceImpl;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.springframework.stereotype.Service;
import org.springframework.web.multipart.MultipartFile;

import java.io.InputStream;

/**
 * <p>
 * 课程科目 服务实现类
 * </p>
 *
 * @author testjava
 * @since 2019-12-25
 */
@Service
public class EduSubjectServiceImpl extends ServiceImpl<EduSubjectMapper, EduSubject> implements EduSubjectService {

    //添加课程分类
    //poi读取excel内容
    @Override
    public void importSubjectData(MultipartFile file) {
        try {

            //1.获取文件输入流
            InputStream inputStream = file.getInputStream();

            //2.创建workbook对象
            Workbook workbook = new HSSFWorkbook(inputStream);

            //3.获取sheet
            Sheet sheet = workbook.getSheetAt(0);

            //int physicalNumberOfRows = sheet.getPhysicalNumberOfRows();
            //获取最后一行的索引值（总共有多少行）
            int lastRowNum = sheet.getLastRowNum();

            //从第二行开始遍历
            for (int i = 1; i <= lastRowNum; i++) {
                //4.获取行
                Row row = sheet.getRow(i);

                //5.获取第一列
                Cell cellOne = row.getCell(0);
                //获取第一列的内容
                String oneCellValue = cellOne.getStringCellValue();

                //添加一级分类
                EduSubject existOneSubject = this.existOneSubject(oneCellValue);
                //没有相同的，则开始添加
                if (existOneSubject == null) {
                    existOneSubject = new EduSubject();
                    //只需要添加，“title”，“parent_id”字段，其他的自动生成
                    existOneSubject.setTitle(oneCellValue);
                    existOneSubject.setParentId("0");
                    baseMapper.insert(existOneSubject);
                }

                //获取一级分类id值
                String pid = existOneSubject.getId();

                //6.获取第二列
                Cell cellTwo = row.getCell(1);
                //获取第二列的内容
                String twoCellValue = cellTwo.getStringCellValue();

                //添加二级分类
                EduSubject existTwoSubject = this.existTwoSubject(twoCellValue, pid);
                if (existTwoSubject == null) {
                    existTwoSubject = new EduSubject();
                    existTwoSubject.setTitle(twoCellValue);
                    existTwoSubject.setParentId(pid);
                    baseMapper.insert(existTwoSubject);
                }
            }

        } catch (Exception e) {
            e.printStackTrace();
        }
    }



    //判断一级分类是否重复
    private EduSubject existOneSubject(String name) {
        //创建QueryWrapper条件封装器对象
        QueryWrapper<EduSubject> wrapper = new QueryWrapper<>();
        //添加条件
        wrapper.eq("title",name);
        wrapper.eq("parent_id","0");
        EduSubject eduSubject = baseMapper.selectOne(wrapper);
        return eduSubject;
    }

    //判断二级分类是否重复
    private EduSubject existTwoSubject(String name,String pid) {
        //创建QueryWrapper条件封装器对象
        QueryWrapper<EduSubject> wrapper = new QueryWrapper<>();
        //添加条件
        wrapper.eq("title",name);
        wrapper.eq("parent_id",pid);
        EduSubject eduSubject = baseMapper.selectOne(wrapper);
        return eduSubject;
    }
}
