package com.example.uploadexcelfileservice.service;

import com.example.uploadexcelfileservice.helper.EmployeeHelper;
import com.example.uploadexcelfileservice.model.Employee;
import com.example.uploadexcelfileservice.repository.EmployeeRepository;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;
import org.springframework.web.multipart.MultipartFile;

import java.io.ByteArrayInputStream;
import java.io.IOException;
import java.util.List;

@Service
public class EmployeeService {

    @Autowired
    private EmployeeRepository repository;

    public void save(MultipartFile file){
        try {
            List<Employee> employees = EmployeeHelper.uploadFileExcel(file.getInputStream());
            repository.saveAll(employees);
        }catch (IOException e){
            throw new RuntimeException("fail to store excel data: " + e.getMessage());
        }
    }

    public List<Employee> getAll(){
        return repository.findAll();
    }

    public ByteArrayInputStream load() {
        List<Employee> employees = repository.findAll();

        ByteArrayInputStream result = EmployeeHelper.downloadFileExcel(employees);
        return result;
    }
}
