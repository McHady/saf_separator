package result;

public enum ResultFileType {
    INDIVIDUAL("individual"),
    ORGANISATION("organisation");

    private final String asString;

    ResultFileType(String asString) {
        this.asString = asString;
    }

    public String asString() {
        return asString;
    }
}
