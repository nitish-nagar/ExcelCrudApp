package com.nitish.excel.excelcrudapp.service;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.Iterator;
import java.util.List;
import java.util.Map;
import java.util.Set;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Service;

import com.nitish.excel.excelcrudapp.entities.User;

@Service
public class UserService {

	// getUserList
	public List<User> getUserList() throws IOException {
		// get excel file
		FileInputStream inputStream = new FileInputStream("C:\\db.xlsx");

		// get workbook and sheet
		XSSFWorkbook workbook = new XSSFWorkbook(inputStream);
		XSSFSheet worksheet = workbook.getSheet("Sheet1");

		List<User> userList = new ArrayList<>();

		// loop through rows in excel file
		Iterator<Row> iterator = worksheet.iterator();
		while (iterator.hasNext()) {
			Row currentRow = iterator.next();

			// skip header
			if (currentRow.getRowNum() == 0) {
				continue;
			}

			// map to User object
			User user = new User();
			Integer id = (int) currentRow.getCell(0).getNumericCellValue();
			user.setId(id);
			user.setName(currentRow.getCell(1).getStringCellValue());
			user.setEmail(currentRow.getCell(2).getStringCellValue());

			userList.add(user);
		}
		inputStream.close();
		workbook.close();
		return userList;
	}

	// getUserById
	public User getUserById(int identifier) throws IOException {
		// get excel file
		FileInputStream inputStream = new FileInputStream("C:\\db.xlsx");

		// get workbook and sheet
		XSSFWorkbook workbook = new XSSFWorkbook(inputStream);
		XSSFSheet worksheet = workbook.getSheet("Sheet1");

		List<User> userList = new ArrayList<>();
		User user = new User();

		// loop through rows in excel file
		Iterator<Row> iterator = worksheet.iterator();
		while (iterator.hasNext()) {
			Row currentRow = iterator.next();

			// skip header
			if (currentRow.getRowNum() == 0) {
				continue;
			}

			// map to User object
			Integer id = (int) currentRow.getCell(0).getNumericCellValue();
			if (id == identifier) {
				user.setId(id);
				user.setName(currentRow.getCell(1).getStringCellValue());
				user.setEmail(currentRow.getCell(2).getStringCellValue());

				userList.add(user);
			}
		}
		inputStream.close();
		workbook.close();
		return user;
	}

	// createUser
	public void createUser(User user) throws IOException {
		// get excel file
		FileInputStream inputStream = new FileInputStream("C:\\db.xlsx");

		// get workbook and sheet
		XSSFWorkbook workbook = new XSSFWorkbook(inputStream);
		XSSFSheet worksheet = workbook.getSheet("Sheet1");
		int rowNum = worksheet.getLastRowNum() + 1;

		Map<Integer, Object[]> data = new HashMap<>();
		data.put(rowNum, new Object[] { rowNum, user.getName().toString(), user.getEmail().toString() });

		Set<Integer> keySet = data.keySet();
		for (Integer key : keySet) {
			Row row = worksheet.createRow(rowNum++);
			Object[] objArr = data.get(key);
			int cellNum = 0;
			for (Object obj : objArr) {
				Cell cell = row.createCell(cellNum++);
				if (obj instanceof String)
					cell.setCellValue((String) obj);
				else if (obj instanceof Integer)
					cell.setCellValue((Integer) obj);
			}
		}
		inputStream.close();
		try {
			FileOutputStream out = new FileOutputStream("C:\\db.xlsx");
			workbook.write(out);
			out.close();
		} catch (Exception e) {
			e.printStackTrace();
		}
		workbook.close();
	}

	// addUserList
	public void addUserList(List<User> userList) throws IOException {
		// get excel file
		FileInputStream inputStream = new FileInputStream("C:\\db.xlsx");

		// get workbook and sheet
		XSSFWorkbook workbook = new XSSFWorkbook(inputStream);
		XSSFSheet worksheet = workbook.getSheet("Sheet1");
		int rowNum = worksheet.getLastRowNum() + 1;

		List<User> userListInner = new ArrayList<>();
		userListInner.addAll(userList);

		for (User u : userListInner) {
			Map<Integer, Object[]> data = new HashMap<>();
			data.put(rowNum, new Object[] { rowNum, u.getName().toString(), u.getEmail().toString() });

			Set<Integer> keySet = data.keySet();
			for (Integer key : keySet) {
				Row row = worksheet.createRow(rowNum++);
				Object[] objArr = data.get(key);
				int cellNum = 0;
				for (Object obj : objArr) {
					Cell cell = row.createCell(cellNum++);
					if (obj instanceof String)
						cell.setCellValue((String) obj);
					else if (obj instanceof Integer)
						cell.setCellValue((Integer) obj);
				}
			}
		}
		inputStream.close();
		try {
			FileOutputStream out = new FileOutputStream("C:\\db.xlsx");
			workbook.write(out);
			out.close();
		} catch (Exception e) {
			e.printStackTrace();
		}
		workbook.close();
	}

	// updateUser
	public void updateUser(int identifier, User user) throws IOException {
		// get excel file
		FileInputStream inputStream = new FileInputStream("C:\\db.xlsx");

		// get workbook and sheet
		XSSFWorkbook workbook = new XSSFWorkbook(inputStream);
		XSSFSheet worksheet = workbook.getSheet("Sheet1");

		// loop through rows in excel file
		Iterator<Row> iterator = worksheet.iterator();
		while (iterator.hasNext()) {
			Row currentRow = iterator.next();

			// skip header
			if (currentRow.getRowNum() == 0) {
				continue;
			}

			// map to User object
			Integer id = (int) currentRow.getCell(0).getNumericCellValue();
			if (id == identifier) {
				currentRow.getCell(0).setCellValue(identifier);
				currentRow.getCell(1).setCellValue(user.getName().toString());
				currentRow.getCell(2).setCellValue(user.getEmail().toString());
			}
		}
		inputStream.close();
		try {
			FileOutputStream out = new FileOutputStream("C:\\db.xlsx");
			workbook.write(out);
			out.close();
		} catch (Exception e) {
			e.printStackTrace();
		}
		workbook.close();
	}

	// modifyUserList
	public void modifyUserList(List<User> userList) throws IOException {
		// get excel file
		FileInputStream inputStream = new FileInputStream("C:\\db.xlsx");

		// get workbook and sheet
		XSSFWorkbook workbook = new XSSFWorkbook(inputStream);
		XSSFSheet worksheet = workbook.getSheet("Sheet1");

		List<User> userListInner = new ArrayList<>();
		userListInner.addAll(userList);

		for (int i = 0; i < userList.size(); i++) {
			int identifier = userList.get(i).getId();
			User user = new User();
			user.setId(identifier);
			user.setName(userList.get(i).getName());
			user.setEmail(userList.get(i).getEmail());

			// loop through rows in excel file
			Iterator<Row> iterator = worksheet.iterator();
			while (iterator.hasNext()) {
				Row currentRow = iterator.next();

				// skip header
				if (currentRow.getRowNum() == 0) {
					continue;
				}

				// map to User object
				Integer id = (int) currentRow.getCell(0).getNumericCellValue();
				if (id == identifier) {
					currentRow.getCell(0).setCellValue(identifier);
					currentRow.getCell(1).setCellValue(user.getName().toString());
					currentRow.getCell(2).setCellValue(user.getEmail().toString());
				}
			}
		}
		inputStream.close();
		try {
			FileOutputStream out = new FileOutputStream("C:\\db.xlsx");
			workbook.write(out);
			out.close();
		} catch (Exception e) {
			e.printStackTrace();
		}
		workbook.close();
	}

	// deleteUser
	public void deleteUser(int identifier) throws IOException {
		// get excel file
		FileInputStream inputStream = new FileInputStream("C:\\db.xlsx");

		// get workbook and sheet
		XSSFWorkbook workbook = new XSSFWorkbook(inputStream);
		XSSFSheet worksheet = workbook.getSheet("Sheet1");

		// loop through rows in excel file
		Iterator<Row> iterator = worksheet.iterator();
		while (iterator.hasNext()) {
			Row currentRow = iterator.next();

			// skip header
			if (currentRow.getRowNum() == 0) {
				continue;
			}

			// map to User object
			Integer id = (int) currentRow.getCell(0).getNumericCellValue();
			if (id == identifier) {
				XSSFRow removingRow = worksheet.getRow(currentRow.getRowNum());
				if (removingRow != null) {
					worksheet.removeRow(removingRow);
				}
			}
		}
		inputStream.close();
		try {
			FileOutputStream out = new FileOutputStream("C:\\db.xlsx");
			workbook.write(out);
			out.close();
		} catch (Exception e) {
			e.printStackTrace();
		}
		workbook.close();
	}
}
