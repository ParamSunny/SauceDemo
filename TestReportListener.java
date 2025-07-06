package utils;

import org.apache.poi.xslf.usermodel.XMLSlideShow;
import org.apache.poi.xslf.usermodel.XSLFShape;
import org.apache.poi.xslf.usermodel.XSLFSlide;
import org.apache.poi.xslf.usermodel.XSLFTextParagraph;
import org.apache.poi.xslf.usermodel.XSLFTextRun;
import org.apache.poi.xslf.usermodel.XSLFTextShape;
import org.testng.ITestContext;
import org.testng.ITestListener;
import org.testng.ITestResult;

import java.awt.*;
import java.io.FileOutputStream;
import java.time.Duration;
import java.util.ArrayList;
import java.util.List;

public class TestReportListener implements ITestListener {
    private static class TestResultInfo {
        String testName;
        String status;
        long durationMillis;

        TestResultInfo(String testName, String status, long durationMillis) {
            this.testName = testName;
            this.status = status;
            this.durationMillis = durationMillis;
        }
    }

    private final List<TestResultInfo> results = new ArrayList<>();
    private long testStartTime;

    @Override
    public void onTestStart(ITestResult result) {
        testStartTime = System.currentTimeMillis();
    }

    @Override
    public void onTestSuccess(ITestResult result) {
        long duration = System.currentTimeMillis() - testStartTime;
        results.add(new TestResultInfo(result.getName(), "PASS", duration));
    }

    @Override
    public void onTestFailure(ITestResult result) {
        long duration = System.currentTimeMillis() - testStartTime;
        results.add(new TestResultInfo(result.getName(), "FAIL", duration));
    }

    @Override
    public void onTestSkipped(ITestResult result) {
        long duration = System.currentTimeMillis() - testStartTime;
        results.add(new TestResultInfo(result.getName(), "SKIPPED", duration));
    }

    @Override
    public void onFinish(ITestContext context) {
        generatePPTReport();
    }

    private void generatePPTReport() {
        try (XMLSlideShow ppt = new XMLSlideShow()) {
            XSLFSlide slide = ppt.createSlide();
            XSLFTextShape title = slide.createTextBox();
            title.setAnchor(new Rectangle(20, 20, 600, 50));
            XSLFTextRun titleRun = title.addNewTextParagraph().addNewTextRun();
            titleRun.setText("Test Execution Report");
            titleRun.setFontSize(24.0);
            titleRun.setBold(true);
            titleRun.setFontColor(new Color(0, 102, 204));

            int yPos = 80;

            // Summary at top
            long total = results.size();
            long passed = results.stream().filter(r -> r.status.equals("PASS")).count();
            long failed = results.stream().filter(r -> r.status.equals("FAIL")).count();
            long skipped = results.stream().filter(r -> r.status.equals("SKIPPED")).count();

            XSLFTextShape summary = slide.createTextBox();
            summary.setAnchor(new Rectangle(20, yPos, 600, 50));
            summary.addNewTextParagraph().addNewTextRun()
                    .setText("Summary: Total = " + total + ", Passed = " + passed + ", Failed = " + failed + ", Skipped = " + skipped);
            yPos += 50;

            // Test results table-like listing
            for (TestResultInfo result : results) {
                XSLFTextShape text = slide.createTextBox();
                text.setAnchor(new Rectangle(20, yPos, 600, 30));
                XSLFTextParagraph para = text.addNewTextParagraph();
                XSLFTextRun run = para.addNewTextRun();
                run.setText(result.testName + " - " + result.status + " (" + result.durationMillis + " ms)");
                run.setFontSize(16.0);
                if (result.status.equals("PASS")) {
                    run.setFontColor(new Color(0, 153, 0));  // Green
                } else if (result.status.equals("FAIL")) {
                    run.setFontColor(Color.RED);
                } else {
                    run.setFontColor(Color.ORANGE);
                }
                yPos += 30;
            }

            try (FileOutputStream out = new FileOutputStream("Test_Report_PPT.pptx")) {
                ppt.write(out);
            }

            System.out.println("Test_Report_PPT.pptx generated successfully!");
        } catch (Exception e) {
            e.printStackTrace();
        }
    }
}
