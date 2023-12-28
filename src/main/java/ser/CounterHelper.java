package ser;

import com.ser.blueline.IDocumentServer;
import com.ser.blueline.IInformationObject;
import com.ser.blueline.ISession;
import com.ser.blueline.metaDataComponents.IStringMatrix;

import java.sql.Date;

public class CounterHelper {

    private static final String GVLObjectNumber = "Agent_ObjectNumber";
    private String descriptor_name;
    private String number_prefix;
    private String counter_keys;
    private String counter_id;
    private String counterStr;
    private ISession session;
    private IDocumentServer documentServer;

    public String getCounterStr() {return counterStr;}
    public CounterHelper(ISession session, IInformationObject object) throws Exception{
        this.session = session;
        this.documentServer = session.getDocumentServer();
        String classID = object.getClassID();
        this.loadConfig(classID);
        this.counterStr = this.returnCounter(classID);
    }
    public CounterHelper(ISession session, String classID) throws Exception{
        this.session = session;
        this.session = session;
        this.documentServer = session.getDocumentServer();
        this.loadConfig(classID);
        this.counterStr = this.returnCounter(classID);
    }


    private void loadConfig(String classID) throws Exception {
        IStringMatrix settingsMatrix = this.documentServer.getStringMatrix(GVLObjectNumber, session);

        for(int i = 0; i < settingsMatrix.getRowCount(); ++i) {
            if (settingsMatrix.getValue(i, 0).equalsIgnoreCase(classID)) {
                this.descriptor_name = settingsMatrix.getValue(i, 1);
                this.number_prefix = settingsMatrix.getValue(i, 2);

                try {
                    this.counter_keys = settingsMatrix.getValue(i, 3);
                } catch (Exception var6) {
                    this.counter_keys = "";
                }

                try {
                    this.counter_id = settingsMatrix.getValue(i, 4);
                } catch (Exception var5) {
                    this.counter_id = "";
                }

                return;
            }
        }

        throw new Exception("Class ID " + classID + " not found in " + "Agent_ObjectNumber");
    }

    private String returnCounter(String classid ) throws Exception {
        Date currentDate = new Date(System.currentTimeMillis());
        String objectNumber = "" + currentDate.toLocalDate();
        String currentYear = objectNumber.split("-")[0];
        if (this.counter_keys.replaceAll(" ", "").isEmpty()) {
            long generatednr = session.getNextCounterValue(3456L, currentYear + classid, 1L, 1L);
            return this.number_prefix + currentYear + "//" + generatednr;
        } else {
            objectNumber = this.number_prefix;
            if (this.counter_id.replaceAll(" ", "").isEmpty()) {
                this.counter_id = "3456";
            }

            int id;
            try {
                id = Integer.parseInt(this.counter_id);
            } catch (Exception var13) {
                throw new Exception("(E) the counter ID, defined in GVL, is not a valid number.");
            }

            String[] keys = this.counter_keys.split(";");
            if (keys.length > 7) {
                throw new Exception("(E) to many custom keys are defined in GVL (max 4).");
            } else {
                for(int i = 0; i < keys.length; ++i) {
                    if (keys[i].equals("YMD")) {
                        objectNumber = objectNumber + ("" + currentDate.toLocalDate()).replaceAll("-", "");
                    }

                    String startValue;
                    String currentDay;
                    if (keys[i].equals("YM")) {
                        startValue = "" + currentDate.toLocalDate();
                        currentDay = startValue.split("-")[1];
                        objectNumber = objectNumber + currentYear + currentDay;
                    }

                    if (keys[i].equals("Y")) {
                        objectNumber = objectNumber + currentYear;
                    }

                    if (keys[i].equals("M")) {
                        startValue = "" + currentDate.toLocalDate();
                        currentDay = startValue.split("-")[1];
                        objectNumber = objectNumber + currentDay;
                    }

                    if (keys[i].equals("D")) {
                        startValue = "" + currentDate.toLocalDate();
                        currentDay = startValue.split("-")[2];
                        objectNumber = objectNumber + currentDay;
                    }

                    if (keys[i].contains("LL")) {
                        startValue = keys[i].replaceAll("LL", "");
                        if (startValue.isEmpty()) {
                            startValue = "1";
                        }

                        int start = Integer.parseInt(startValue);
                        long generatednr = session.getNextCounterValue((long)id, currentYear + classid, (long)start, 1L);
                        objectNumber = objectNumber + generatednr;
                    }

                    if (keys[i].contains("CC")) {
                        objectNumber = objectNumber + keys[i].replaceAll("CC", "");
                    }



                }

                return objectNumber;
            }
        }
    }


}
