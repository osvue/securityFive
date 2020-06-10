package org.javaboy.withjpa.dao;

import org.javaboy.withjpa.model.User;
import org.springframework.data.jpa.repository.JpaRepository;

/**/
public interface UserDao extends JpaRepository<User,Long> {
    User findUserByUsername(String username);
}
