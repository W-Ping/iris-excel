package com.iris.excelfile.data;

import com.iris.excelfile.model.*;
import org.apache.commons.lang3.StringUtils;

import java.math.BigDecimal;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.util.*;

/**
 * @author liu_wp
 * @date Created in 2019/3/7 16:20
 * @see
 */
public class TestData {

    private final static String[] DCC_GROUP_NAME = {"A组DCC", "B组DCC"};
    private final static String[] SC_GROUP_NAME = {"A组SC", "B组DSC"};
    private static DateTimeFormatter format = DateTimeFormatter.ofPattern("yyyy-MM-dd_HHmmss");

    public static String createUniqueFileName(String fileName) {
        LocalDateTime now = LocalDateTime.now();
        String format = now.format(TestData.format);
        if (StringUtils.isNotBlank(fileName)) {
            format = format.concat("_").concat(fileName);
        }
        return format;
    }


    public static List<List<String>> createHeadList() {
        List<List<String>> head = new ArrayList<>();
        List<String> head1 = new ArrayList<>();
        head1.add("表1");
        head1.add("表1");
        head1.add("表21");
        List<String> head2 = new ArrayList<>();
        head2.add("表2");
        head2.add("表21");
        List<String> head3 = new ArrayList<>();
        head3.add("表3");
        head3.add("表3");
        head3.add("表3");
        List<String> head4 = new ArrayList<>();
        head4.add("表4");
        head4.add("表4");
        head4.add("表4");
        head.add(head1);
        head.add(head2);
        head.add(head3);
        head.add(head4);
        return head;
    }

    public static List<ExportTestModel> createTestListJavaMode() {
        List<ExportTestModel> model1s = new ArrayList<ExportTestModel>();
        for (int i = 0; i < 5; i++) {
            ExportTestModel model1 = new ExportTestModel();
            model1.setP1("第一列，第" + (i + 1) + "行");
            model1.setP2("32323JJfdf");
            model1.setP3(33 + i);
            model1.setP4(44);
            model1.setP5("55ces");
            model1.setP6(666.2f);
            model1.setP7(new BigDecimal("23.13991399"));
            model1.setP8(new Date());
            model1.setP9("PPPP9999");
            model1.setP10(1111.77 + i);
            model1s.add(model1);
        }
        return model1s;
    }

    public static List<ExportTestModel> createTestListJavaModeStyle() {
        List<ExportTestModel> model1s = new ArrayList<ExportTestModel>();
        for (int i = 0; i < 5; i++) {
            ExportTestModel model1 = new ExportTestModel();
            model1.setP1("第一列，第" + (i + 1) + "行");
            model1.setP2("32323JJfdf");
            model1.setP3(33 + i);
            model1.setP4(44);
            model1.setP5("55ces");
            model1.setP6(666.2f);
            model1.setP7(new BigDecimal("23.13991399"));
            model1.setP8(new Date());
            model1.setP9("PPPP9999");
            model1.setP10(1111.77 + i);
//            model1.setCellStyleMap(cellStyleMap);
            model1s.add(model1);
        }
        return model1s;
    }

    public static List<ExcelTest2Model> createTestListJavaMode2() {
        List<ExcelTest2Model> model1s = new ArrayList<>();
        for (int i = 0; i < 3; i++) {
            ExcelTest2Model model1 = new ExcelTest2Model();
            model1.setP1("第一列，第" + (i + 1) + "行");
            model1.setP2("32323JJfdf");
            model1.setP3(22 + i);
            model1.setP4(88);
            model1.setP5("66ces");
            model1.setP6(888.2f);
            model1.setP7(new BigDecimal("45.13991399"));
            model1.setP8(new Date());
            model1.setP9(331.77 + i);
            model1s.add(model1);
        }
        return model1s;
    }

    public static List<ExportTest3Model> createTestListJavaMode3() {
        List<ExportTest3Model> model1s = new ArrayList<>();
        for (int i = 0; i < 3; i++) {
            ExportTest3Model model1 = new ExportTest3Model();
            model1.setP1("第一列，第" + (i + 1) + "行");
            model1.setP2("39087623JJfdf");
            model1.setP3(42 + i);
            model1.setP4(912);
            model1.setP5("66ces");
            model1.setP6(333.2f);
            model1.setP7(new BigDecimal("45.13991399"));
            model1s.add(model1);
        }
        return model1s;
    }


    /**
     * @param groupName
     * @return
     */
    public static CrossReportGroupNameModel mockGroupCrossReportRate(String groupName) {
        CrossReportGroupNameModel crossReportGroupNameModel = new CrossReportGroupNameModel();
        crossReportGroupNameModel.setEmpCode(groupName);
        return crossReportGroupNameModel;
    }

    /**
     * @param groupName
     * @return
     */
    public static CrossReportRateGroupNameModel mockGroupSCCrossReportRate(String groupName) {
        CrossReportRateGroupNameModel crossReportGroupNameModel = new CrossReportRateGroupNameModel();
        crossReportGroupNameModel.setEmpCode(groupName);
        return crossReportGroupNameModel;
    }


    /**
     * @param size
     * @return
     */
    public static Map<String, List<CrossReportModel>> mockDCCCrossReportData(int size) {
        Map<String, List<CrossReportModel>> result = new LinkedHashMap<>();
        for (int i = 0; i < DCC_GROUP_NAME.length; i++) {
            result.put(DCC_GROUP_NAME[i], mockDCCCrossReportData(DCC_GROUP_NAME[i], size));
        }
        return result;
    }

    /**
     * @param size
     * @return
     */
    public static Map<String, List<CrossReportModel>> mockSCCrossReportData(int size) {
        Map<String, List<CrossReportModel>> result = new LinkedHashMap<>();
        for (int i = 0; i < SC_GROUP_NAME.length; i++) {
            result.put(SC_GROUP_NAME[i], mockSCCrossReportData(SC_GROUP_NAME[i], size));
        }
        return result;
    }

    /**
     * @param size
     * @return
     */
    public static Map<String, List<CrossReportRateModel>> mockSCCrossReportRateData(int size) {
        Map<String, List<CrossReportRateModel>> result = new LinkedHashMap<>();
        for (int i = 0; i < SC_GROUP_NAME.length; i++) {
            result.put(SC_GROUP_NAME[i], mockCrossReportRateData(SC_GROUP_NAME[i], size));
        }
        return result;
    }

    /**
     * SC 数据
     *
     * @param groupName
     * @param size
     * @return
     */
    public static List<CrossReportModel> mockSCCrossReportData(String groupName, int size) {
        List<CrossReportModel> model1s = new ArrayList<>();
        for (int i = 0; i < size; i++) {
            CrossReportModel model1 = new CrossReportModel();
            model1.setEmpType(groupName);
            model1.setDateVersion(LocalDate.now().toString());
            model1.setCustomerFlowAllTargetValue(10);
            model1.setEmpCode(groupName + "_" + i);
            model1.setCustomerFlowTargetValue(10);
            model1.setCustomerFlowAllActualValue(10);
            model1.setCustomerFlowActualValue(10);
            model1.setIncomingCallAllTargetValue(null);
            model1.setIncomingCallAllActualValue(null);
            model1.setIncomingCallTargetValue(7);
            model1.setIncomingCallActualValue(8);
            model1.setNetworkAllTargetValue(null);
            model1.setNetworkAllActualValue(null);
            model1.setNetworkTargetValue(11);
            model1.setNetworkActualValue(12);
            model1.setInitiativeCollectAllTargetValue(null);
            model1.setInitiativeCollectAllActualValue(null);
            model1.setInitiativeCollectTargetValue(15);
            model1.setInitiativeCollectActualValue(16);
            model1.setRecommendAllTargetValue(10);
            model1.setRecommendAllActualValue(10);
            model1.setRecommendTargetValue(10);
            model1.setRecommendActualValue(10);
            model1.setBuyAgainAllTargetValue(10);
            model1.setBuyAgainAllActualValue(10);
            model1.setBuyAgainTargetValue(10);
            model1.setBuyAgainActualValue(10);
            model1.setActiveAllTargetValue(10);
            model1.setActiveAllActualValue(10);
            model1.setActiveTargetValue(10);
            model1.setActiveActualValue(10);
            model1.setDormantAllTargetValue(10);
            model1.setDormantAllActualValue(10);
            model1.setDormantTargetValue(10);
            model1.setDormantActualValue(10);

            model1.setCustomerFlowTotalTargetValue(null);
            model1.setCustomerFlowTotalActualValue(null);
            model1.setIntoStoreSaleTotalTargetValue(null);
            model1.setIntoStoreSaleTotalActualValue(null);
            model1.setQuotedPriceCountTargetValue(10);
            model1.setQuotedPriceCountActualValue(10);
            model1.setOrderCountTargetValue(10);
            model1.setOrderCountActualValue(10);
            model1.setInvoicedCountTargetValue(10);
            model1.setInvoicedCountActualValue(10);
            model1.setFirstTestDriveCountTargetValue(10);
            model1.setFirstTestDriveCountActualValue(10);
            model1.setFinancialCountTargetValue(10);
            model1.setFinancialCountActualValue(10);
            model1.setInsuranceCountTargetValue(10);
            model1.setInsuranceCountActualValue(10);
            model1.setSkuGoodsAmountTargetValue(new BigDecimal(29000.99));
            model1.setSkuGoodsAmountActualValue(new BigDecimal(19000.99));
            model1.setExtendWarrantyCountTargetValue(10);
            model1.setExtendWarrantyCountActualValue(10);
            model1.setOtherCountTargetValue(10);
            model1.setOtherCountActualValue(10);
            model1s.add(model1);

        }
        return model1s;
    }


    /**
     * DCC 数据
     *
     * @param groupName
     * @param size
     * @return
     */
    public static List<CrossReportModel> mockDCCCrossReportData(String groupName, int size) {
        List<CrossReportModel> model1s = new ArrayList<>();
        for (int i = 0; i < size; i++) {
            CrossReportModel model1 = new CrossReportModel();
            model1.setEmpType(groupName);
            model1.setDateVersion(LocalDate.now().toString());
            model1.setCustomerFlowAllTargetValue(8);
            model1.setEmpCode(groupName + "_" + i);
            model1.setCustomerFlowTargetValue(8);
            model1.setCustomerFlowAllActualValue(8);
            model1.setCustomerFlowActualValue(8);
            model1.setIncomingCallAllTargetValue(8);
            model1.setIncomingCallAllActualValue(8);
            model1.setIncomingCallTargetValue(8);
            model1.setIncomingCallActualValue(8);
            model1.setNetworkAllTargetValue(8);
            model1.setNetworkAllActualValue(8);
            model1.setNetworkTargetValue(8);
            model1.setNetworkActualValue(8);
            model1.setInitiativeCollectAllTargetValue(8);
            model1.setInitiativeCollectAllActualValue(8);
            model1.setInitiativeCollectTargetValue(8);
            model1.setInitiativeCollectActualValue(8);
            model1.setRecommendAllTargetValue(8);
            model1.setRecommendAllActualValue(8);
            model1.setRecommendTargetValue(8);
            model1.setRecommendActualValue(8);
            model1.setBuyAgainAllTargetValue(8);
            model1.setBuyAgainAllActualValue(8);
            model1.setBuyAgainTargetValue(8);
            model1.setBuyAgainActualValue(8);
            model1.setActiveAllTargetValue(8);
            model1.setActiveAllActualValue(8);
            model1.setActiveTargetValue(8);
            model1.setActiveActualValue(8);
            model1.setDormantAllTargetValue(8);
            model1.setDormantAllActualValue(8);
            model1.setDormantTargetValue(8);
            model1.setDormantActualValue(8);

            model1.setCustomerFlowTotalTargetValue(null);
            model1.setCustomerFlowTotalActualValue(null);
            model1.setIntoStoreSaleTotalTargetValue(null);
            model1.setIntoStoreSaleTotalActualValue(null);
            model1.setQuotedPriceCountTargetValue(8);
            model1.setQuotedPriceCountActualValue(8);
            model1.setOrderCountTargetValue(8);
            model1.setOrderCountActualValue(8);
            model1.setInvoicedCountTargetValue(8);
            model1.setInvoicedCountActualValue(8);
            model1.setFirstTestDriveCountTargetValue(8);
            model1.setFirstTestDriveCountActualValue(8);
            model1.setFinancialCountActualValue(null);
            model1.setInsuranceCountTargetValue(null);
            model1.setInsuranceCountActualValue(null);
            model1.setSkuGoodsAmountTargetValue(null);
            model1.setSkuGoodsAmountActualValue(null);
            model1.setExtendWarrantyCountTargetValue(null);
            model1.setExtendWarrantyCountActualValue(null);
            model1.setOtherCountTargetValue(null);
            model1.setOtherCountActualValue(null);
            model1s.add(model1);
        }
        return model1s;
    }

    public static CrossReportGroupModel2 mockCrossReportTemplate() {
        CrossReportGroupModel2 model1 = new CrossReportGroupModel2();
        int type = (int) (1 + Math.random() * (100 - 10 + 1));
        model1.setEmpType("1");
        model1.setDateVersion(LocalDate.now().toString());
        model1.setCustomerFlowAllTargetValue(199 + type);
        model1.setCustomerFlowAllActualValue(10 * type + 100);
        model1.setInitiativeCollectAllTargetValue(299);
        model1.setInitiativeCollectAllActualValue(398);
        model1.setCustomerFlowTargetValue(340 * type);
        model1.setCustomerFlowActualValue(340);
        model1.setIncomingCallAllTargetValue(340);
        model1.setIncomingCallAllActualValue(340);
        model1.setIncomingCallTargetValue(340);
        model1.setIncomingCallActualValue(340 + type);
        model1.setNetworkAllTargetValue(340);
        model1.setNetworkAllActualValue(340);
        model1.setNetworkTargetValue(340);
        model1.setNetworkActualValue(340);

        model1.setInitiativeCollectTargetValue(340);
        model1.setInitiativeCollectActualValue(340);
        model1.setRecommendAllTargetValue(340);
        model1.setRecommendAllActualValue(340);
        model1.setRecommendTargetValue(340);
        model1.setRecommendActualValue(340);
        model1.setBuyAgainAllTargetValue(340);
        model1.setBuyAgainAllActualValue(340);
        model1.setBuyAgainTargetValue(340);
        model1.setBuyAgainActualValue(340);
        model1.setActiveAllTargetValue(340 + type);
        model1.setActiveAllActualValue(340);
        model1.setActiveTargetValue(340);
        model1.setActiveActualValue(340);
        model1.setDormantAllTargetValue(340);
        model1.setDormantAllActualValue(340 + type);
        model1.setDormantTargetValue(340);
        model1.setDormantActualValue(340);
        model1.setCustomerFlowTotalTargetValue(340 + type);
        model1.setCustomerFlowTotalActualValue(340);
        model1.setIntoStoreSaleTotalTargetValue(340);
        model1.setIntoStoreSaleTotalActualValue(340);
        model1.setQuotedPriceCountTargetValue(340);
        model1.setQuotedPriceCountActualValue(340);
        model1.setOrderCountTargetValue(340);
        model1.setOrderCountActualValue(340);
        model1.setInvoicedCountTargetValue(340);
        model1.setInvoicedCountActualValue(340);
        model1.setFirstTestDriveCountTargetValue(340);
        model1.setFirstTestDriveCountActualValue(340);

//        model1.setFinancialCountTargetValue(11);
//        model1.setFinancialCountActualValue(12);
//        model1.setInsuranceCountTargetValue(22);
//        model1.setInsuranceCountActualValue(23);
        return model1;
    }

    /**
     * @param size
     * @return
     */
    public static List<CrossReportGroupModel> mockCrossReportTemplate(int size, String nameType) {
        List<CrossReportGroupModel> model1s = new ArrayList<>();
        for (int i = 0; i < size; i++) {
            int type = (int) (1 + Math.random() * (100 - 10 + 1));
            CrossReportGroupModel model1 = new CrossReportGroupModel();
            model1.setEmpType("1");
            model1.setDateVersion(LocalDate.now().toString());
            model1.setCustomerFlowAllTargetValue(199 + type);
            model1.setCustomerFlowAllActualValue(10 * type + 100);
            model1.setInitiativeCollectAllTargetValue(299);
            model1.setInitiativeCollectAllActualValue(398);
            if (!"market".equals(nameType)) {
                model1.setCustomerFlowTargetValue(340 * type);
                model1.setCustomerFlowActualValue(340);
                model1.setIncomingCallAllTargetValue(340);
                model1.setIncomingCallAllActualValue(340);
                model1.setIncomingCallTargetValue(340);
                model1.setIncomingCallActualValue(340 + type);
                model1.setNetworkAllTargetValue(340);
                model1.setNetworkAllActualValue(340);
                model1.setNetworkTargetValue(340);
                model1.setNetworkActualValue(340);

                model1.setInitiativeCollectTargetValue(340);
                model1.setInitiativeCollectActualValue(340);
                model1.setRecommendAllTargetValue(340);
                model1.setRecommendAllActualValue(340);
                model1.setRecommendTargetValue(340);
                model1.setRecommendActualValue(340);
                model1.setBuyAgainAllTargetValue(340);
                model1.setBuyAgainAllActualValue(340);
                model1.setBuyAgainTargetValue(340);
                model1.setBuyAgainActualValue(340);
                model1.setActiveAllTargetValue(340 + type);
                model1.setActiveAllActualValue(340);
                model1.setActiveTargetValue(340);
                model1.setActiveActualValue(340);
                model1.setDormantAllTargetValue(340);
                model1.setDormantAllActualValue(340 + type);
                model1.setDormantTargetValue(340);
                model1.setDormantActualValue(340);
                model1.setCustomerFlowTotalTargetValue(340 + type);
                model1.setCustomerFlowTotalActualValue(340);
                model1.setIntoStoreSaleTotalTargetValue(340);
                model1.setIntoStoreSaleTotalActualValue(340);
                model1.setQuotedPriceCountTargetValue(340);
                model1.setQuotedPriceCountActualValue(340);
                model1.setOrderCountTargetValue(340 * type);
                model1.setOrderCountActualValue(340);
                model1.setInvoicedCountTargetValue(340);
                model1.setInvoicedCountActualValue(340);
                model1.setFirstTestDriveCountTargetValue(340);
                model1.setFirstTestDriveCountActualValue(340);
                if (i == 0) {
                    model1.setFinancialCountTargetValue(340);
                    model1.setFinancialCountActualValue(340);
                    model1.setInsuranceCountTargetValue(340);
                    model1.setInsuranceCountActualValue(340);
                    model1.setSkuGoodsAmountTargetValue(new BigDecimal("23888.16945678"));
                    model1.setSkuGoodsAmountActualValue(new BigDecimal("23888.16345678"));
                    model1.setExtendWarrantyCountTargetValue(345);
                    model1.setExtendWarrantyCountActualValue(345);
                    model1.setOtherCountTargetValue(345);
                    model1.setOtherCountActualValue(345 + type);
                }
            }
            model1s.add(model1);

        }
        return model1s;
    }

    /**
     * @param empType
     * @param size
     * @return
     */
    public static List<CrossReportRateModel> mockCrossReportRateData(String empType, int size) {
        List<CrossReportRateModel> model1s = new ArrayList<>();
        for (int i = 0; i < size; i++) {
            CrossReportRateModel model1 = new CrossReportRateModel();
            model1.setEmpCode(empType + "_" + i);
            model1s.add(model1);
//            model1.setPhone("1381899999");
//            model1.setDate(new Date());
//            model1.setDate2("2013-01-21 15:10:20");
//            model1.setPhone2(1881899999);
        }
        return model1s;
    }
}
