package com.luwei.module.easyexcel.service;

import com.luwei.module.easyexcel.pojo.User;
import lombok.extern.slf4j.Slf4j;
import org.springframework.stereotype.Service;

import java.util.List;

/**
 * @author wukun
 * @since 2019/10/20
 */
@Service
@Slf4j
public class UserService {
    /**
     * 模拟批量插入用户
     *
     * @param users userList
     */
    public void saveBatchUser(List<User> users) {

        log.info("/*------- UserService#saveBatchUser() : {} -------*/", users);
    }
}
