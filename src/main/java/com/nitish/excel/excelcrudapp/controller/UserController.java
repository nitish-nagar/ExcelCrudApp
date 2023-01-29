package com.nitish.excel.excelcrudapp.controller;

import java.io.IOException;
import java.util.List;

import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.http.HttpStatus;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.DeleteMapping;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.PathVariable;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.PutMapping;
import org.springframework.web.bind.annotation.RequestBody;
import org.springframework.web.bind.annotation.RestController;

import com.nitish.excel.excelcrudapp.entities.User;
import com.nitish.excel.excelcrudapp.service.UserService;

@RestController
public class UserController {

	@Autowired
	private UserService userService;

	// getUserList
	@GetMapping(value = "/users", produces = MediaType.APPLICATION_JSON_VALUE)
	public ResponseEntity<List<User>> getUserList() throws IOException {
		List<User> userList = userService.getUserList();
		return new ResponseEntity<>(userList, HttpStatus.OK);
	}

	// getUserById
	@GetMapping(value = "/users/{id}", produces = MediaType.APPLICATION_JSON_VALUE)
	public ResponseEntity<User> getUserById(@PathVariable int id) throws IOException {
		User user = userService.getUserById(id);
		return new ResponseEntity<>(user, HttpStatus.OK);
	}

	// createUser
	@PostMapping(value = "/users", produces = MediaType.APPLICATION_JSON_VALUE)
	public ResponseEntity<User> createUser(@RequestBody User user) throws IOException {
		userService.createUser(user);
		return new ResponseEntity<>(user, HttpStatus.OK);
	}

	// addUserList
	@PostMapping(value = "/userList", produces = MediaType.APPLICATION_JSON_VALUE)
	public ResponseEntity<List<User>> addUserList(@RequestBody List<User> userList) throws IOException {
		userService.addUserList(userList);
		return new ResponseEntity<>(userList, HttpStatus.OK);
	}

	// updateUser
	@PutMapping(value = "/users/{id}", produces = MediaType.APPLICATION_JSON_VALUE)
	public ResponseEntity<User> updateUser(@RequestBody User user, @PathVariable int id) throws IOException {
		userService.updateUser(id, user);
		return new ResponseEntity<>(user, HttpStatus.OK);
	}

	// modifyUserList
	@PutMapping(value = "/userList", produces = MediaType.APPLICATION_JSON_VALUE)
	public ResponseEntity<List<User>> modifyUserList(@RequestBody List<User> userList) throws IOException {
		userService.modifyUserList(userList);
		return new ResponseEntity<>(userList, HttpStatus.OK);
	}

	// deleteUser
	@DeleteMapping(value = "/users/{id}")
	public ResponseEntity<User> deleteUser(@PathVariable int id) throws IOException {
		userService.deleteUser(id);
		return new ResponseEntity<>(HttpStatus.OK);
	}
}
