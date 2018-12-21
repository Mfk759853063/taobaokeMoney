package com.vbn.taobaokeMoney.controller;

import com.vbn.taobaokeMoney.excelOperation.OperationExcel;
import org.apache.logging.log4j.Logger;
import org.apache.poi.sl.draw.geom.Path;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.*;
import org.springframework.web.multipart.MultipartFile;
import org.springframework.web.servlet.ModelAndView;
import org.springframework.web.servlet.mvc.support.RedirectAttributes;

import java.nio.file.*;
import java.util.Map;


/**
 * Created by vbn on 2018/12/21.
 */

@Controller
public class simpleController {


    @RequestMapping(value = "/v1/index")
    private ModelAndView uploadExcel(ModelAndView mv) {
        mv.setViewName("upload");
        return mv;
    }

    // 上传业务处理
    @PostMapping("/v1/upload")
    private ModelAndView singleFileUpload(ModelAndView mv,@RequestParam("file") MultipartFile file , RedirectAttributes redirectAttributes) {
        System.out.print("上传文件了***********  *******");
        mv.setViewName("result");
        if (file.isEmpty()) {
            mv.setViewName("error");
            mv.addObject("message", "请选择一个文件");
            return mv;
        }
        try {
            byte[] bytes = file.getBytes();
            java.nio.file.Path path = Paths.get("upload/"+file.getOriginalFilename());
            Files.write(path, bytes);
            Map resultMap =  OperationExcel.readExcel(path.toString());
            if (resultMap==null) {
                mv.setViewName("error");
                mv.addObject("message", "出错啦");
            } else {
                mv.addObject("message", "计算成功");
                mv.addObject("map", resultMap);
            }
        } catch (Exception e) {
            mv.setViewName("error");
            mv.addObject("message", e.getMessage());
            return mv;
        }
        return mv;
    }
}
