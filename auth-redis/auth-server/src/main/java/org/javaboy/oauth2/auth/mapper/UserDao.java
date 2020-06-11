package org.javaboy.oauth2.auth.mapper;


import org.javaboy.oauth2.auth.domain.User;
import org.springframework.data.jpa.repository.JpaRepository;

/**/
public interface UserDao extends JpaRepository<User,Long> {
    User findUserByUsername(String username);
}
