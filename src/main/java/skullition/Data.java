package skullition;

public class Data {
    private final String reportType;
    private final String personName;
    private final String email;
    private final String period;
    private final String courseName;
    private final String coursePrice;
    private final String transactionAmount;
    private final String totalPaymentGateway;
    private final String cutPaymentGateway;
    private final String administration;
    private final String incomePercentage;
    private final String incomeBeforeTax;
    private final String taxPercentage;
    private final String taxAmount;
    private final String endIncome;

    public Data(String reportType, String personName, String email, String period, String courseName, String coursePrice, String transactionAmount, String totalPaymentGateway, String cutPaymentGateway, String administration, String incomePercentage, String incomeBeforeTax, String taxPercentage, String taxAmount, String endIncome) {
        this.reportType = reportType;
        this.personName = personName;
        this.email = email;
        this.period = period;
        this.courseName = courseName;
        this.coursePrice = coursePrice;
        this.transactionAmount = transactionAmount;
        this.totalPaymentGateway = totalPaymentGateway;
        this.cutPaymentGateway = cutPaymentGateway;
        this.administration = administration;
        this.incomePercentage = incomePercentage;
        this.incomeBeforeTax = incomeBeforeTax;
        this.taxPercentage = taxPercentage;
        this.taxAmount = taxAmount;
        this.endIncome = endIncome;
    }

    public String getReportType() {
        return reportType;
    }

    public String getPersonName() {
        return personName;
    }

    public String getEmail() {
        return email;
    }

    public String getPeriod() {
        return period;
    }

    public String getCourseName() {
        return courseName;
    }

    public String getCoursePrice() {
        return coursePrice;
    }

    public String getTransactionAmount() {
        return transactionAmount;
    }

    public String getTotalPaymentGateway() {
        return totalPaymentGateway;
    }

    public String getCutPaymentGateway() {
        return cutPaymentGateway;
    }

    public String getAdministration() {
        return administration;
    }

    public String getIncomePercentage() {
        return incomePercentage;
    }

    public String getIncomeBeforeTax() {
        return incomeBeforeTax;
    }

    public String getTaxPercentage() {
        return taxPercentage;
    }

    public String getTaxAmount() {
        return taxAmount;
    }

    public String getEndIncome() {
        return endIncome;
    }
}
