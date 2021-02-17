import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.Arrays;
import java.util.Objects;
import java.util.Properties;
import java.util.concurrent.Executors;
import java.util.stream.Collectors;

public class Main {

    private static String[] getCompanyNameParts(String configPath, String delimiter) {

        try (var stream = new FileInputStream(configPath)) {

            var properties = new Properties();
            properties.load(stream);

            var nameString = properties.getProperty("includingInCompanyFile", "");
            return nameString.split(delimiter);

        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }

        return null;

    }

    private static final String listNameCriteria = "list.xlsx";
    private static final String root = ".";

    public static void main(String[] args) {

        var companyNameParts = getCompanyNameParts("settings.ini", ";");

        //getting list files
        var listFiles = Arrays.stream(Objects.requireNonNull(new File(root).listFiles()))
                .filter(file -> file.isFile() && file.getName().contains(listNameCriteria))
                .map(file -> root + "/" + file.getName())
                .collect(Collectors.toList());

        var service = Executors.newFixedThreadPool(listFiles.size());

        for (var file : listFiles) {

            var future = service.submit(new SeparateProcess(companyNameParts, root, file));
        }

        service.shutdown();
    }
}
