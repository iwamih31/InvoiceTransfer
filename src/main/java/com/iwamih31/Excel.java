package com.iwamih31;

import java.io.InputStream;
import java.io.OutputStream;
import java.io.UnsupportedEncodingException;
import java.net.URLEncoder;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.util.ArrayList;
import java.util.List;

import javax.servlet.http.HttpServletResponse;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Row.MissingCellPolicy;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.springframework.web.multipart.MultipartFile;

public class Excel {

	MissingCellPolicy cellPolicy = MissingCellPolicy.CREATE_NULL_AS_BLANK;

	public String transfer_Excel(
			MultipartFile output_model,
			MultipartFile input_model,
			MultipartFile input_data,
			HttpServletResponse httpServletResponse) {
		___console_Out___("Excel#transfer_Excel(…)開始");
    String message = "データの移行が";
    String add_Message = "";
    int remaining = 0;
		try (
				// 出力側モデル
				InputStream outModel_InputStream = output_model.getInputStream();
				Workbook outModel_Book = WorkbookFactory.create(outModel_InputStream);
				// 入力側モデル
				InputStream inModel_InputStream = input_model.getInputStream();
				Workbook inModel_Book = WorkbookFactory.create(inModel_InputStream);
				// 入力側データ
    		InputStream data_InputStream = input_data.getInputStream();
				Workbook data_Book = WorkbookFactory.create(data_InputStream);
				// ダウンロード用ストリーム
				OutputStream outputStream = httpServletResponse.getOutputStream()
			) {
			___console_Out___("try開始");

			// 出力用シートを指定
			Sheet outModel_Sheet = outModel_Book.getSheetAt(0);
			// 入力側モデルシートを指定
			Sheet inModel_Sheet = inModel_Book.getSheetAt(0);
			// 入力側データシートを指定
			Sheet data_Sheet = data_Book.getSheetAt(0);

			// 出力用シートの行数を取得
			int outModel_Row_Size = outModel_Sheet.getLastRowNum() + 1;
			___console_Out___("row_Size = " + outModel_Row_Size);
			// 出力用シートの行数分ループ
			for (int i = 0; i < outModel_Row_Size; i++) {
				___console_Out___("　i = " + i);
				// 行を取得
				Row row = outModel_Sheet.getRow(i);
				// 列数を取得
				int outModel_Column_Size = 0;
				if(row != null) outModel_Column_Size = row.getLastCellNum() + 1;
				___console_Out___("column_Size = " + outModel_Column_Size);
				// 列数分ループ
				for (int j = 0; j < outModel_Column_Size; j++) {
					___console_Out___("j = " + j);
					// チェックするセルの値を取得
					String check_Value = get_Value(outModel_Sheet, i, j, cellPolicy);
					___console_Out___("check_Value = " + check_Value);
					// セルの値が$$～$$の形に当てはまる場合
					if (check_Value.startsWith("$$") && check_Value.endsWith("$$")) {
						___console_Out___("対象セルです");

						// 対象セルに対応する data_Sheet 側セルのリストを作成
						List<Cell> data_Cell_List = data_Cell_List(inModel_Sheet, data_Sheet, check_Value);
						int data_Cell_List_size = data_Cell_List.size();
						___console_Out___("data_Cell_List_size = " + data_Cell_List_size);
						// リストの要素数分ループ
						for (int k = 0; k < data_Cell_List.size(); k++) {
							___console_Out___("k = " + k);
							int out_Row = i + k;
							___console_Out___("out_Row = " + out_Row);
							int out_Column = j;
							// 出力側セルの値を取得
							String outModel_Cell_Value = get_Value(outModel_Sheet, out_Row, out_Column, cellPolicy);
							___console_Out___("outModel_Cell_Value = " + outModel_Cell_Value);
							String outModel_Next_Cell_Value = "$$$$";
							// 最後の行で無ければ
							if (out_Row < outModel_Row_Size - 1) {
								// その1つ下のセルの値を取得
								outModel_Next_Cell_Value = get_Value(outModel_Sheet, out_Row + 1, out_Column, cellPolicy);
								___console_Out___("outModel_Next_Cell_Value = " + outModel_Next_Cell_Value);
							}
							// 出力側セルを取得
							Cell outModel_Cell = get_Cell(outModel_Sheet, out_Row, out_Column);
							// データ側セルを取得
							Cell data_Cell = data_Cell_List.get(k);
							// 出力側セルにデータ側セルの値を入力
							set_Cell_Value(outModel_Cell, data_Cell);
							// 終了条件を満たした場合
							if (is_last(outModel_Cell_Value, outModel_Next_Cell_Value)) {
								// 残数を追加
								remaining += k;
								// ループを抜ける
								break;
							}
						}
					___console_Out___("対象セル終了");
					}
				}
			}

			___console_Out___("未処理データ数 = " + remaining);
			if (remaining > 0) add_Message += remaining + " 件のデータが未処理です";
			// ファイル名を取得
			String file_Name = output_model.getResource().getFilename().replace(".xlsx", "");
			___console_Out___("file_Name = " + file_Name);
			// ファイル名を指定して保存
			if (response_Making(httpServletResponse, file_Name)) {
				// outputStream に 出力用ブックを書き込み
				outModel_Book.write(outputStream);
				message += "完了しました " + add_Message;
			}
    } catch (Exception e) {
    	message += "正常に行われませんでした";
    	e.printStackTrace();
    }
		___console_Out___("Excel#transfer_Excel(…)終了");
		return message;
	}


	private List<Cell> data_Cell_List(Sheet inModel_Sheet, Sheet data_Sheet, String check_Value) {
		// data_Cell 格納用リスト作成
		List<Cell> data_Cell_List = new ArrayList<>();
		// 入力側モデルの同じ値のセルと同じ位置の入力側データのセルを取得
		Cell target_Cell = get_Like_Cell(inModel_Sheet, check_Value);
		// データのセルが在れば
		if (target_Cell != null) {
			___console_Out___("データのリストを作成します");
			// そのセルの行番号を取得
			int target_Row = target_Cell.getRowIndex();
			___console_Out___("target_Row = " + target_Row);
			// そのセルの列番号を取得
			int target_Column = target_Cell.getColumnIndex();
			___console_Out___("target_Column = " + target_Column);
			// データシートの行数を取得
			int data_Row_Size = data_Sheet.getLastRowNum() + 1;
			___console_Out___("data_Row_Size = " + data_Row_Size);
			// 行数分ループ
			for (target_Row = target_Row + 0; target_Row < data_Row_Size; target_Row++) {
				___console_Out___("target_Row = " + target_Row);
				Cell data_Cell = get_Cell(data_Sheet, target_Row, target_Column);
				// リストにセルをセット
				data_Cell_List.add(data_Cell);
				/* 終了条件チェック */
				// テンプレート側の同位置にあるセルの値を取得
				String check_Cell_Value = get_Value(inModel_Sheet, target_Row, target_Column, cellPolicy);
				___console_Out___("check_Cell_Value = " + check_Cell_Value);
				String next_Cell_Value = "$$$$";
				// inModel_Sheet の最後の行以外なら
				if (target_Row < data_Row_Size - 1) {
					//その 1つ下のセルの値を取得
					next_Cell_Value = get_Value(inModel_Sheet, target_Row + 1, target_Column, cellPolicy);
					___console_Out___("next_Cell_Value = " + next_Cell_Value);
				}
				// 終了条件を満たした場合ループを抜ける
				if (is_last(check_Cell_Value, next_Cell_Value)) break;
			}
		}
		return data_Cell_List;
	}


	private Cell get_Cell(Sheet sheet, int row_Number, int column_Number) {
		Cell cell = null;
		Row row = sheet.getRow(row_Number);
		if(row != null) cell = row.getCell(column_Number, cellPolicy);
		return cell;
	}


	private String get_Value(Sheet sheet, int row, int column, MissingCellPolicy cellPolicy) {
		String cell_Value = "";
		// sheet の指定位置にあるセルの値を取得
		Row inModel_Row = sheet.getRow(row);
		if (inModel_Row != null) {
			Cell cell = inModel_Row.getCell(column, cellPolicy);
			cell_Value = get_Value(cell);
		}
		return cell_Value;
	}


	private String set_Cell_Value(Cell set_Cell, Cell value_Cell) {
		String message = "value_Cell = null";
		if (value_Cell == null) {
			set_Cell.setCellValue("null");
		} else {
			switch (value_Cell.getCellType()) {
				case STRING:	//文字型
					set_Cell.setCellValue(value_Cell.getStringCellValue());
					break;
				case NUMERIC: // 数値型、日付型 ※日付もNUMERICと判定される
					set_Cell.setCellValue(value_Cell.getNumericCellValue());
					break;
				case FORMULA: // Excel関数型 ※例）NOW()、SUM()などのExcelの関数
					switch (value_Cell.getCachedFormulaResultType()) {
					case BLANK:
						set_Cell.setCellValue("");
						break;
					case BOOLEAN:
						String booLean = "false";
						if(value_Cell.getBooleanCellValue() == true) booLean = "true";
						set_Cell.setCellValue(booLean);
						break;
					case ERROR:
						set_Cell.setCellValue("エラー！");
						break;
					case FORMULA:
						set_Cell.setCellValue(value_Cell.getCellFormula());
						break;
					case NUMERIC:
						set_Cell.setCellValue(value_Cell.getNumericCellValue());
						break;
					case STRING:
						set_Cell.setCellValue(value_Cell.getStringCellValue());
						break;
					case _NONE:
						set_Cell.setCellValue("");
						break;
					default:
						break;
					}
					break;
				case BOOLEAN: // 真偽型 ※例）TRUE、FALSE
					String booLean = "false";
					if(value_Cell.getBooleanCellValue() == true) booLean = "true";
					set_Cell.setCellValue(booLean);
					break;
				case BLANK: // 空 ※セルに値がセットされていない場合の型
					 set_Cell.setCellValue("");
					break;
				case _NONE:
					set_Cell.setCellValue("");
					break;
				default:
					set_Cell.setCellValue("?");
					break;
			}
			message = "set_Cell_Value = " + get_Value(set_Cell);
		}
		___console_Out___(message);
		return message;
	}


	private String get_Value(Cell cell) {
		String value = "null";
		if (cell != null) {
			switch(cell.getCellType()) {
				case STRING:	//文字型
					value = cell.getStringCellValue();
					break;
				case NUMERIC: //	数値型、日付型※日付もNUMERICと判定されます
					value = String.valueOf(cell.getNumericCellValue());
					break;
				case FORMULA: // Excel関数型 ※例）NOW()、SUM()などのExcelの関数
					switch (cell.getCachedFormulaResultType()) {
						case BLANK:
							value = "";
							break;
						case BOOLEAN:
							String booLean = "false";
							if(cell.getBooleanCellValue() == true) booLean = "true";
							value = booLean;
							break;
						case ERROR:
							value = "エラー！";
							break;
						case FORMULA:
							value = cell.getCellFormula();
							break;
						case NUMERIC:
							value = String.valueOf(cell.getNumericCellValue());
							break;
						case STRING:
							value = cell.getStringCellValue();
							break;
						case _NONE:
							value = "";
							break;
						default:
							value = "?";
							break;
					}
					break;
				case BLANK: // 空 ※セルに値がセットされていない場合の型
					value = "";
					break;
				case _NONE:
					value = "";
					break;
				default:
					value = "?";
					break;
			}
		}
		return value;
	}


	private Cell get_Data_Cell(Sheet inModel_sheet, Sheet data_sheet, Cell outModel_Cell) {

		Cell data_Cell = null;
		// 入力用シートの行数を取得
		int row_Size = inModel_sheet.getLastRowNum() + 1;
		___console_Out___("row_Size = " + row_Size);
		// 出力用シートの行数分ループ
		for (int i = 0; i < row_Size; i++) {
			// 行を取得
			Row row = inModel_sheet.getRow(i);
			// 列数を取得
			int column_Size = row.getLastCellNum() + 1;
			___console_Out___("column_Size = " + column_Size);
			// 列数分ループ
			for (int j = 0; j < column_Size; j++) {
				// セルを取得
				Cell inModel_Cell = get_Cell(inModel_sheet, i, j);
				// セルの値を比較して同じ値が在れば
				if(get_Value(inModel_Cell) == get_Value(outModel_Cell)) {
					// data_sheet の同位置の値を data_Cell に代入して
					data_Cell = data_sheet.getRow(i).getCell(j, cellPolicy);
					// data_Cell を返してメソッド終了
					return data_Cell;
				}
			}
		}
		return data_Cell;
	}


	private Cell get_Like_Cell(Sheet sheet, Cell check_Cell) {
		Cell like_cell = null;
		// 入力用シートの行数を取得
		int row_Size = sheet.getLastRowNum() + 1;
		___console_Out___("row_Size = " + row_Size);
		// 出力用シートの行数分ループ
		for (int i = 0; i < row_Size; i++) {
			// 行を取得
			Row row = sheet.getRow(i);
			// 列数を取得
			int column_Size = row.getLastCellNum() + 1;
			___console_Out___("column_Size = " + column_Size);
			// 列数分ループ
			for (int j = 0; j < column_Size; j++) {
				// セルを取得
				Cell cell = get_Cell(sheet, i, j);
				// セルの値を比較して同じ値が在れば
				if(get_Value(cell) == get_Value(check_Cell)) {
					// data_sheet の同位置の値を data_Cell に代入して
					like_cell = cell;
					___console_Out___("like_cell = " + like_cell);
					// data_Cell を返してメソッド終了
					return like_cell;
				}
			}
		}
		return like_cell;
	}

	private Cell get_Like_Cell(Sheet sheet, String check_Value) {
		Cell like_cell = null;
		// シートの行数を取得
		int row_Size = sheet.getLastRowNum() + 1;
		___console_Out___("row_Size = " + row_Size);
		// 行数分ループ
		for (int i = 0; i < row_Size; i++) {
			// 行を取得
			Row row = sheet.getRow(i);
			// 列数を取得
			int column_Size = row.getLastCellNum() + 1;
			___console_Out___("column_Size = " + column_Size);
			// 列数分ループ
			for (int j = 0; j < column_Size; j++) {
				// セルを取得
				Cell cell = get_Cell(sheet, i, j);
				// セルの値を比較して同じ値ならば
				if(check_Value.equals(get_Value(cell))) {
					// セルを like_cell に代入して
					like_cell = cell;
					___console_Out___("like_cell = " + like_cell);
					// like_cell を返してメソッド終了
					return like_cell;
				}
			}
		}
		return like_cell;
	}


	private List<String[]> get_Data(Sheet sheet) {
		// 行数を取得
		int rows = sheet.getLastRowNum() + 1;
		___console_Out___("rows = " + rows);
		// 列数を取得
		int cols = sheet.getRow(0).getLastCellNum() + 1;
		___console_Out___("cols = " + cols);
		// 行データ配列のリストを作成
		List<String[]> data_List= new ArrayList<>();
		// 行数分ループ
		for (int i = 0; i < rows; i++) {
			// 行のセルの値を入れる配列を作成
			String[] rowData = new String[cols];
			// 列数分ループ
			for (int j = 0; j < cols; j++) {
				// セルの値を取得して行データ配列に格納
				Cell cell = get_Cell(sheet, i, j);
				String value = "";
				if (cell != null) value = get_Value(cell);
				rowData[j] = value;
				___console_Out___("rowData[" + j + "] = " + rowData[j]);
			}
			// 行データ配列のリストに行データ配列を追加
			data_List.add(rowData);
			___console_Out___("row" + i + "終了");
		}
		return data_List;
	}

	// output_Sheet の check_Value と同じ値のセルから下を value_List の値に書き換える
	private int transfer_Data(Sheet out_Sheet, String check_Value, List<String> value_List) {
		___console_Out___("transfer_Data(Sheet output_Sheet, String check_Value, List<String> value_List)開始");
		// 残りのデータ数
		int remaining = 0;
		// 出力用シートの行数を取得
		int output_Rows = out_Sheet.getLastRowNum() + 1;
		___console_Out___("output_Rows = " + output_Rows);
		// データを取得
		for (int i = 0; i < output_Rows; i++) {
			// 行を指定
	    Row row = out_Sheet.getRow(i);
			// 出力用シートの列数を取得
	    int output_Cols = 0;
	    if (row != null) {
	    	output_Cols = row.getLastCellNum() + 1;
	    	___console_Out___("output_Cols = " + output_Cols);
			}
	    for (int j = 0; j < output_Cols; j++) {
	    	// セルを指定
        Cell cell = row.getCell(j);
        // セルの値を取得
        String value = get_Value(cell);

        ___console_Out___("value = " + value);


        // check_Valueと同じ値のセルが在れば
        if (value.equals(check_Value)) {
        	// value_List の要素数分ループ
        	for (int k = 0; k < value_List.size(); k++) {
        		int row_Num = k + i;
        		// 対象セル取得
        		Cell deta_set_Cell = out_Sheet.getRow(row_Num).getCell(j);
        		// 対象セルの現在の値を取得
        		String cell_Value = get_Value(deta_set_Cell);
            ___console_Out___("output_Sheet.cell_Value = " + cell_Value);
        		// 現在の値 を value_List の要素で書き換え

        		deta_set_Cell.setCellValue(value_List.get(k));
        		___console_Out___(cell_Value + " を " + get_Value(deta_set_Cell) + " に書き換えました");
        		// 対象セルの1つ下のセルの値を取得
            String next_Cell_Value = "$$$$";
      			if (row_Num + 1 < output_Rows) next_Cell_Value = get_Value(out_Sheet.getRow(row_Num + 1).getCell(j));
        		___console_Out___("output_Sheet.next_Cell_Value = " + next_Cell_Value);
        		// 終了条件を満たした場合
        		if (is_last(cell_Value, next_Cell_Value)) {
      				// 残りのデータ数をセット
      				remaining = value_List.size() - 1 - k;
      				// ループを抜ける
      				break;
      			}
					}
	      }
	    }
		}
		___console_Out___("transfer_Data()．remaining = " + remaining);

		// 残りのデータ数を返す
		return remaining;
	}

	private boolean is_last(String cell_Value, String next_Cell_Value) {
		boolean is_last = false;
		// cell_Value に%%～%%マークが在ればループ終了
		if (cell_Value.startsWith("%%") && cell_Value.endsWith("%%")) is_last = true;
		// cell_Value に$$$～$$$マークが在ればループ終了
		if (cell_Value.startsWith("$$$") && cell_Value.endsWith("$$$"))  is_last = true;
		// cell_Value の下のセルに$$～$$マークが在ればループ終了
		if (next_Cell_Value.startsWith("$$") && next_Cell_Value.endsWith("$$"))  is_last = true;

		___console_Out___("is_last = " + is_last);

		return is_last;
	}

	private boolean response_Making(HttpServletResponse response, String file_Name) {
		boolean is_Make = false;
		// ダウンロードファイルのファイルネームを作成
		file_Name = with_Now(file_Name) + ".xlsx";
		___console_Out___("file_Name = " + file_Name);
		String encodedFilename = with_Now("create") + ".xlsx";
		___console_Out___("file_Name を " + encodedFilename + "に設定しました");
		try {
			// 日本語のファイル名が使える様に変換
			encodedFilename = URLEncoder.encode(file_Name, "UTF-8");
		} catch (UnsupportedEncodingException e) {
			e.printStackTrace();
		}
		___console_Out___("file_Name を " + encodedFilename + "に設定しました");
		response.setContentType("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
	  response.setHeader("Content-Disposition", "attachment;filename=\"" + encodedFilename + "\"");
	  response.setCharacterEncoding("UTF-8");
	  is_Make = true;
		return is_Make;
	}

	private String with_Now(String head_String) {
		String now = now().replaceAll("[^0-9]", ""); // 現在日時の数字以外を "" に変換
//	String now = now().replaceAll("[^\\d]", "");  ←こちらでもOK
		now = now.substring(0, now.length()-3); // 後ろから3文字を取り除く
		return head_String + now;
	}

	public String now() {
		// 現在日時を取得
		LocalDateTime now = LocalDateTime.now();
		// 表示形式を指定
		DateTimeFormatter dateTimeFormatter = DateTimeFormatter.ofPattern("yyyy/MM/dd HH:mm:ss.SSS");
		return dateTimeFormatter.format(now);
	}

	/** コンソールに String を出力 */
	public static void ___console_Out___(String message) {
		System.out.println(message);
		System.out.println("*");
	}
}
