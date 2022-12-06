package PdfUpdate;

import java.io.BufferedReader;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.IOException;
import java.io.InputStream;
import java.io.InputStreamReader;
import java.io.PrintWriter;
import java.io.StringWriter;
import java.io.Writer;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.util.ArrayList;
import java.util.List;
import java.util.Scanner;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import org.w3c.tidy.Tidy;

import com.aspose.pdf.DocSaveOptions;
import com.aspose.pdf.Document;
import com.aspose.words.DocumentBuilder;
import com.aspose.words.FindReplaceOptions;
import com.aspose.words.Font;
import com.aspose.words.IReplacingCallback;
import com.aspose.words.ImportFormatMode;
import com.aspose.words.Node;
import com.aspose.words.NodeCollection;
import com.aspose.words.ParagraphFormat;
import com.aspose.words.ReplaceAction;
import com.aspose.words.ReplacingArgs;
import com.aspose.words.SaveFormat;
import com.aspose.words.TxtSaveOptions;

public class Main {

	public static void main(String[] args) throws Exception {

		Scanner sc = new Scanner(System.in);
		System.out.println("Enter PDF path");
		String Path = sc.next();
		System.out.println("Path = " + Path);


		BufferedReader bi = new BufferedReader(new InputStreamReader(System.in));

		ConvertPDFtoWord();

		System.out.println("=================");
		com.aspose.words.Document doc = new com.aspose.words.Document("2015 D&O DemoCo User2.docx");
		TxtSaveOptions opts = new TxtSaveOptions();
		doc.save("out.txt", opts);

		String text = new String(Files.readAllBytes(Paths.get("out.txt")), StandardCharsets.UTF_8);
		Pattern p = Pattern.compile("\\{(.*?)\\}");
		Matcher m = p.matcher(text);

		List variables = new ArrayList();
		while (m.find()) {
			variables.add(m.group());
		}
		
		if (text.contains("{Name}")) {
			format();
		}

	
		String option = "Y";
		while (option.equalsIgnoreCase("Y")) {
				System.out.println("Enter field");
				String Field = "{" + sc.next() + "}";
				System.out.println("Enter value");
				String Value = bi.readLine();
				System.out.println(Field + "=>" + Value);
				variables.remove(Field);
				replace(Field, Value, false);
				System.out.println("More field (Y/N) :");
				option = sc.next();
		}

		if (variables.size() != 0) {
			for (int i = 0; i < variables.size(); i++) {
				replace((String) variables.get(i), "{variablenotprovided}", true);
			}
		}

		removeEmptyParagraph();

//		Eroor
		if (var.size() != 0) {
			com.aspose.words.Document doc1 = new com.aspose.words.Document();
			DocumentBuilder builder1 = new DocumentBuilder(doc1);
			Font font = builder1.getFont();
			font.setName("Arial");
			font.setSize(8);
			for (int i = 0; i < var.size(); i++) {
				builder1.write(var.get(i) + " = Variable Not Found");
				builder1.insertParagraph();
			}
			doc1.save("Error.docx");

//		    mergigng
			com.aspose.words.Document doc2 = new com.aspose.words.Document("2015 D&O DemoCo User2.docx");
			com.aspose.words.Document doc3 = new com.aspose.words.Document("Error.docx");
			doc2.appendDocument(doc3, ImportFormatMode.KEEP_SOURCE_FORMATTING);
			doc2.save("merge.docx");

		}
		if (var.size() != 0) {
			convertDocToPDF("merge.docx");
		} else {
			convertDocToPDF("2015 D&O DemoCo User2.docx");
		}

		File f1 = new File("Error.docx");
		if (f1.exists()) {
			f1.delete();
		}

		File f2 = new File("merge.docx");
		if (f2.exists()) {
			f2.delete();
		}
		File f3 = new File("2015 D&O DemoCo User2.docx");
		if (f3.exists()) {
			f3.delete();
		}
		File f4 = new File("out.txt");
		if (f4.exists()) {
			f4.delete();
		}

		bi.close();
		System.out.println("=================");
		System.out.println("Done");

	}

	static void traverse() throws Exception {
		com.aspose.words.Document doc = new com.aspose.words.Document("2015 D&O DemoCo User2.docx");
		Node node = doc;
		while (node != doc.getLastSection().getBody().getLastParagraph().getLastChild()
				&& node.nextPreOrder(doc) != null) {
			node = node.nextPreOrder(doc);

			System.out.println(Node.nodeTypeToString(node.getNodeType()));

			System.out.println(node.getText());

		}
	}

	static void removeEmptyParagraph() throws Exception {
		com.aspose.words.Document doc = new com.aspose.words.Document("2015 D&O DemoCo User2.docx");
		Node node = doc;
		Node n1;
		while (node != doc.getLastSection().getBody().getLastParagraph().getLastChild()
				&& node.nextPreOrder(doc) != null) {
			node = node.nextPreOrder(doc);

			if (Node.nodeTypeToString(node.getNodeType()).equals("Paragraph")) {
				if (node.getText().length() == 1) {
					n1 = node.previousPreOrder(doc);
					node.remove();
					node = n1;
				}
			}

		}
		doc.save("2015 D&O DemoCo User2.docx", SaveFormat.DOCX);
	}

	static class nameFormat implements IReplacingCallback {
		public int replacing(ReplacingArgs e) throws Exception {
			Node currentNode = e.getMatchNode();
			DocumentBuilder builder = new DocumentBuilder((com.aspose.words.Document) e.getMatchNode().getDocument());
			builder.moveTo(currentNode);
			ParagraphFormat p = builder.getParagraphFormat();
			double leftIndent = p.getLeftIndent();
			double spaceBefore = p.getSpaceBefore();
			if (currentNode.getText().equals("{Name}")) {
				leftIndent = p.getLeftIndent() + p.getFirstLineIndent();
				p.setFirstLineIndent(0);
				p.clearFormatting();
				p.setSpaceBefore(spaceBefore);

				p.setLineSpacing(14.1);
				p.setLineSpacingRule(1);

			} else {
				p.clearFormatting();
				spaceBefore = 0.0;
			}
			p.setSpaceAfter(0.0);

			p.setLeftIndent(leftIndent);

			return ReplaceAction.SKIP;
		}
	}

	static class ReplaceWithHtmlEvaluator implements IReplacingCallback {

		String Value, Field;
		boolean Error;

		public ReplaceWithHtmlEvaluator(String field, String value, boolean error) {
			Value = value;
			Field = field;
			Error = error;
		}

		public int replacing(ReplacingArgs e) throws Exception {

			// This is a Run node that contains either the beginning or the complete match.
			Node currentNode = e.getMatchNode();
			// create Document Buidler and insert MergeField
			DocumentBuilder builder = new DocumentBuilder((com.aspose.words.Document) e.getMatchNode().getDocument());

			builder.moveTo(currentNode);

			if (!currentNode.getText().equals(Field)) {

				NodeCollection nod = e.getMatchNode().getParentNode().getChildNodes();

				if (!Error) {
					builder.insertHtml(Value);
				}

				boolean flag = false;

				for (int i = 0; i < nod.getCount(); i++) {
					Node n9 = nod.get(i);

					if (n9.getText().contains("{")) {
						flag = true;
						nod.removeAt(i);
						i--;
					} else if (n9.getText().contains("}")) {
						nod.removeAt(i);
						i--;
						flag = false;
					} else if (flag == true && !n9.getText().contains(Value)) {
						nod.removeAt(i);
						i--;
					} else {

					}
				}
				if (Error) {
					builder.write(Value);
				}
			} else {

				currentNode.remove();
				if (Error) {
					Font f = builder.getFont();
					f.setBold(false);
					f.setSize(f.getSize() - 2);
					builder.write(Value);
				} else {
					builder.insertHtml(Value);
				}
			}
			return ReplaceAction.SKIP;
		}
	}

	public static void format() throws Exception {

		com.aspose.words.Document doc = new com.aspose.words.Document("2015 D&O DemoCo User2.docx");
		DocumentBuilder builder = new DocumentBuilder(doc);
		FindReplaceOptions findReplaceOptions = new FindReplaceOptions();
		findReplaceOptions.setReplacingCallback(new nameFormat());

		doc.getRange().replace("{Name}", "{Name}&p");
		doc.getRange().replace("{Name}", "", findReplaceOptions);
		doc.getRange().replace("{Address}", "", findReplaceOptions);

		doc.save("2015 D&O DemoCo User2.docx", SaveFormat.DOCX);
	}

	static List var = new ArrayList();

	public static String HTMLCheck(String html) throws IOException {
		String returnHTML = html;
		Tidy tidy = new Tidy();
		InputStream in = new ByteArrayInputStream(html.getBytes("UTF-8"));
		ByteArrayOutputStream errorOutputStream = new java.io.ByteArrayOutputStream();
		PrintWriter errorPrintWriter = new java.io.PrintWriter(errorOutputStream, true); // second param enables
																							// autoflush so you don't
																							// have to manually flush
																							// the printWriter
		Writer stringWriter = new StringWriter();
		tidy.setQuiet(true);
		tidy.setMakeClean(true);
		tidy.setShowWarnings(false);
		tidy.setTidyMark(false);
		tidy.setXHTML(true);
		tidy.setXmlTags(false);
		tidy.setErrout(errorPrintWriter);
		org.w3c.tidy.Node parsedNode = tidy.parse(in, stringWriter);

		String s = errorOutputStream.toString();

		if (s.contains("Error")) {
			int ns = s.indexOf("Error");
			int last = s.indexOf("\n", ns);
			String l = s.substring(ns, last);
			returnHTML = l;
		}
		in.close();
		errorOutputStream.close();
		errorPrintWriter.close();
		stringWriter.close();

		return returnHTML;
	}

	public static void replace(String field, String value, boolean Flag) throws Exception {

		com.aspose.words.Document doc = new com.aspose.words.Document("2015 D&O DemoCo User2.docx");
		DocumentBuilder builder = new DocumentBuilder(doc);

		if (Flag) {
			doc.getRange().replace(field, field + "- NOT FOUND");
			doc.save("2015 D&O DemoCo User2.docx", SaveFormat.DOCX);
		} else {
			String text = new String(Files.readAllBytes(Paths.get("out.txt")), StandardCharsets.UTF_8);
			boolean flag = false;
			if (text.contains(field)) {
				String newValue = HTMLCheck(value);
				if (value.compareTo(newValue) != 0) {
					flag = true;
					value = newValue;
				}

				FindReplaceOptions findReplaceOptions = new FindReplaceOptions();
				findReplaceOptions.setReplacingCallback(new ReplaceWithHtmlEvaluator(field, value, flag));

				doc.getRange().replace(field, "", findReplaceOptions);

				doc.save("2015 D&O DemoCo User2.docx", SaveFormat.DOCX);
			} else {
				var.add(field);
			}
		}

	}

	public static void ConvertPDFtoWord() throws Exception {
		Document doc = new Document("variableSample.pdf");
		DocSaveOptions saveOptions = new DocSaveOptions();
		saveOptions.setMode(DocSaveOptions.RecognitionMode.Flow);
		saveOptions.setFormat(DocSaveOptions.DocFormat.DocX);
		doc.save("2015 D&O DemoCo User2.docx", saveOptions);
		doc.close();

	}

	public static void convertDocToPDF(String file) {
		com.aspose.words.Document doc;
		try {
			doc = new com.aspose.words.Document(file);
			doc.save("replace.pdf");
		} catch (Exception e) {
			e.printStackTrace();
		}

	}

}
