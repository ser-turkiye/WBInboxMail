package ser;

import com.ser.blueline.*;
import com.ser.blueline.bpm.*;
import com.ser.blueline.metaDataComponents.IStringMatrix;
import de.ser.doxis4.agentserver.AgentExecutionResult;
import de.ser.doxis4.agentserver.UnifiedAgent;
import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;
import org.json.JSONArray;
import org.json.JSONObject;

import java.text.SimpleDateFormat;
import java.util.*;
import java.util.concurrent.TimeUnit;

import static ser.Utils.loadTableRows;
import static ser.Utils.saveWBInboxExcel;

public class EscalationMailLoad extends UnifiedAgent {
    private Logger log = LogManager.getLogger();
    IDocument mailTemplate = null;
    JSONObject projects = new JSONObject();
    JSONObject tasks = new JSONObject();
    List<ITask> subprcss = new ArrayList<>();
    List<ITask> mainprcss = new ArrayList<>();
    ProcessHelper helper;

    @Override
    protected Object execute() {
        if (getBpm() == null)
            return resultError("Null BPM object");

        com.spire.license.LicenseProvider.setLicenseKey(Conf.Licences.SPIRE_XLS);

        Utils.session = getSes();
        Utils.bpm = getBpm();
        Utils.server = Utils.session.getDocumentServer();
        Utils.loadDirectory(Conf.WBInboxMailPaths.EscalationMailPaths);

        try {
            JSONObject mcfg = Utils.getMailConfig();

            helper = new ProcessHelper(Utils.session);
            projects = Utils.getProjectWorkspaces(helper);
            mailTemplate = null;

            for(String prjn : projects.keySet()){
                IInformationObject prjt = (IInformationObject) projects.get(prjn);
                IDocument dtpl = Utils.getTemplateDocument(prjt, Conf.MailTemplates.Reviewer);
                if(dtpl == null){continue;}
                mailTemplate = dtpl;
                break;
            }

            if(mailTemplate == null){throw new Exception("Mail template not found.");}

            notifyForReviewer(mcfg,mailTemplate);
            notifyForConsalidators(mcfg,mailTemplate);
            notifyForDccs(mcfg,mailTemplate);

            log.info("Tested.");
        } catch (Exception e) {
            //throw new RuntimeException(e);
            log.error("Exception       : " + e.getMessage());
            log.error("    Class       : " + e.getClass());
            log.error("    Stack-Trace : " + e.getStackTrace() );
            return resultRestart("Exception : " + e.getMessage(),10);
        }

        log.info("Finished");
        return resultSuccess("Ended successfully");
    }
    private IInformationObject getProject(String prjn){
        if(projects.has(prjn)){return (IInformationObject) projects.get(prjn);}
        if(projects.has("!" + prjn)){return null;}

        IInformationObject iprj = Utils.getProjectWorkspace(prjn, helper);
        if(iprj == null){
            projects.put("!" + prjn, "[[ " + prjn + " ]]");
            return null;
        }
        projects.put(prjn, iprj);
        return iprj;
    }
    private void notifyForReviewer(JSONObject mcfg, IDocument mtpl) throws Exception {
        JSONObject MailList = new JSONObject();
        JSONObject ConsalidatorList = new JSONObject();
        JSONObject DCCList = new JSONObject();
        IWorkbasket wbsk = null;
        String wbMail = "", mailExcelPath = "", uniqueId = "";
        JSONObject mbms = new JSONObject();
        int listSize = 0;

        subprcss = Utils.getSubReviewProcesses(helper,projects);
        for(ITask task : subprcss){
            log.info("    -> start ");
            int tcnt = 0;

            tcnt++;
            log.info(" *** subprocess task [" + tcnt + "] : " + task.getDisplayName());

            String tsid = task.getID();
            log.info(" *** subprocess id [" + tcnt + "] : " + tsid);

            String clid = task.getClassID();
            log.info(" *** subprocess class id [" + tcnt + "] : " + clid);
            if(!clid.equals(Conf.ClassIDs.SubReview)){continue;}

            IProcessInstance proi = task.getProcessInstance();
            if(proi == null){continue;}
            log.info(" *** subprocess proi [" + tcnt + "] : " + proi.getDisplayName());

            IDocument wdoc = (IDocument) proi.getMainInformationObject();
            if(wdoc == null){continue;}
            log.info(" *** subprocess wdoc [" + tcnt + "] : " + wdoc.getDisplayName());

            String prjn = wdoc.getDescriptorValue(Conf.Descriptors.ProjectNo, String.class);
            if(prjn == null || prjn.isEmpty()){continue;}
            log.info(" *** subprocess prjn [" + tcnt + "] : " + prjn);

            IDocument prjCardDoc = getProjectCard(prjn);
            List<String> cnslList = consolidatorList((IInformationObject) task);
            String dccMailList = (prjCardDoc != null ? dccMails(prjCardDoc) : "");
            String prjDocRvwDrtS = (prjCardDoc != null ? prjCardDoc.getDescriptorValue("ccmPRJCard_ReviewerDrtn") : "0");
            String isAutoComplete = (prjCardDoc != null ? prjCardDoc.getDescriptorValue("ObjectState") : "false");
            long prjRvwDrtn = Long.parseLong(prjDocRvwDrtS);

            log.info("    -> class-name : " + task.getName());
            log.info("    -> class-id.doc : " + wdoc.getClassID());
            log.info("    -> class-id.task : " + clid);
            log.info("    -> display : " + wdoc.getDisplayName());

            Date tbgn = null, tend = new Date();
            if(task.getReadyDate() != null){
                tbgn = task.getReadyDate();
            }

            log.info("    -> begin date : " + tbgn + "   -> end date : " + tend);

            long durd  = 0L;
            double durh  = 0.0;
            if(tend != null && tbgn != null) {
                long diff = (tend.getTime() > tbgn.getTime() ? tend.getTime() - tbgn.getTime() : tbgn.getTime() - tend.getTime());
                durd = TimeUnit.DAYS.convert(diff, TimeUnit.MILLISECONDS);

                long diffSeconds = diff / 1000;
                long diffMinutes = diff / (60 * 1000);
                long diffHours = diff / (60 * 60 * 1000);

                log.info("    -> duration : " + durd + "   -> prj duration : " + prjRvwDrtn);

                if(durd<=prjRvwDrtn){continue;}
                durh = ((TimeUnit.MINUTES.convert(diff, TimeUnit.MILLISECONDS) - (durd * 24 * 60)) * 100 / 60) / 100d;
            }

            if(!Objects.equals(isAutoComplete, "false") && isAutoComplete!=null) {
                moveCurrentTaskToNext(task);
                log.info("Task is ["+ task.getID() +"] auto completed. Duration:" + durd);
                continue;
            }
            ///escalation list creating for consalidators
            for(String cnslId : cnslList) {
                if(cnslId.isEmpty()){continue;}
                IUser mmbr = getDocumentServer().getUserByLoginName(getSes() , getUserLoginByWB(cnslId));
                if(mmbr == null){continue;}
                ConsalidatorList.accumulate(mmbr.getEMailAddress(), task.getID());
            }
            ///escalation list creating for dcc
            DCCList.accumulate(dccMailList, task.getID());
        }

        for(String keyStr : ConsalidatorList.keySet()) {
            wbMail = keyStr;
            Object value = ConsalidatorList.get(keyStr);
            uniqueId = UUID.randomUUID().toString();
            mailExcelPath = Utils.exportDocument(mtpl, Conf.WBInboxMailPaths.EscalationMailPaths, Conf.MailTemplates.Project + "[" + uniqueId + "]");
            int dcnt = 0;
            if (value instanceof JSONArray) {
                JSONArray arr = (JSONArray) value;
                listSize = arr.length();
                loadTableRows(mailExcelPath, Conf.WBInboxMailSheetIndex.Mail, "Task", Conf.WBInboxMailRowGroups.MailColInx, listSize);

                mbms.put("Fullname", keyStr + " (Consolidators)");
                mbms.put("Count", listSize + "");

                for(int i = 0; i < arr.length(); i++){
                    String tID = arr.getString(i);
                    dcnt++;
                    ITask xtsk = getBpm().findTask(tID);
                    if(xtsk == null){continue;}
                    IProcessInstance proi = xtsk.getProcessInstance();
                    IInformationObject mdoc = (proi != null ? proi.getMainInformationObject() : null);

                    IWorkbasket orwb = xtsk.getOriginalWorkbasket();
                    String owbn = "";
                    //if(orwb != null && orwb.getID() != wbsk.getID()){
                        //owbn = orwb.getFullName();
                    //}

                    Date tbgn = null, tend = new Date();
                    if(xtsk.getReadyDate() != null){
                        tbgn = xtsk.getReadyDate();
                    }

                    long durd  = 0L;
                    double durh  = 0.0;
                    if(tend != null && tbgn != null) {
                        long diff = (tend.getTime() > tbgn.getTime() ? tend.getTime() - tbgn.getTime() : tbgn.getTime() - tend.getTime());
                        durd = TimeUnit.DAYS.convert(diff, TimeUnit.MILLISECONDS);
                        durh = ((TimeUnit.MINUTES.convert(diff, TimeUnit.MILLISECONDS) - (durd * 24 * 60)) * 100 / 60) / 100d;
                    }

                    String rcvf = "", rcvo = "";
                    if(xtsk.getPreviousWorkbasket() != null){
                        rcvf = xtsk.getPreviousWorkbasket().getFullName();
                    }
                    if(tbgn != null){
                        rcvo = (new SimpleDateFormat("dd-MM-yyyy HH:mm")).format(tbgn);
                    }

                    String prjn = "",  mdno = "", mdrn = "", mdnm = "";
                    if(mdoc != null &&  Utils.hasDescriptor((IInformationObject) mdoc, Conf.Descriptors.ProjectNo)){
                        prjn = mdoc.getDescriptorValue(Conf.Descriptors.ProjectNo, String.class);
                    }
                    if(mdoc != null &&  Utils.hasDescriptor((IInformationObject) mdoc, Conf.Descriptors.DocNumber)){
                        mdno = mdoc.getDescriptorValue(Conf.Descriptors.DocNumber, String.class);
                    }
                    if(mdoc != null &&  Utils.hasDescriptor((IInformationObject) mdoc, Conf.Descriptors.Revision)){
                        mdrn = mdoc.getDescriptorValue(Conf.Descriptors.Revision, String.class);
                    }
                    if(mdoc != null &&  Utils.hasDescriptor((IInformationObject) mdoc, Conf.Descriptors.Name)){
                        mdnm = mdoc.getDescriptorValue(Conf.Descriptors.Name, String.class);
                    }

                    mbms.put("Task" + dcnt, xtsk.getName() + " (" + xtsk.getDescriptorValue("LayerName") + ")");
                    mbms.put("ProcessTitle" + dcnt, (proi != null ? proi.getDisplayName() : ""));
                    mbms.put("Delegation" + dcnt, (owbn != null ? owbn : ""));
                    mbms.put("ProjectNo" + dcnt, (prjn != null  ? prjn : ""));
                    mbms.put("DocNo" + dcnt, (mdno != null  ? mdno : ""));
                    mbms.put("RevNo" + dcnt, (mdrn != null  ? mdrn : ""));
                    mbms.put("DocName" + dcnt, (mdnm != null  ? mdnm : ""));
                    mbms.put("ReceivedFrom" + dcnt, (rcvf != null ? rcvf : ""));
                    mbms.put("ReceivedOn" + dcnt, (rcvo != null ? rcvo : ""));
                    mbms.put("DurDay" + dcnt, durd + "");
                    mbms.put("DurHour" + dcnt, durh + "");
                    mbms.put("DoxisLink" + dcnt, mcfg.get("webBase") + helper.getTaskURL(xtsk.getID()));
                    mbms.put("DoxisLink" + dcnt + ".Text", "( Link )");
                }
            }
            else if (value instanceof String) {
                String tID = (String) value;

                mbms.put("Fullname", keyStr + "( Consalidator)");
                mbms.put("Count", 1 + "");

                loadTableRows(mailExcelPath, Conf.WBInboxMailSheetIndex.Mail, "Task", Conf.WBInboxMailRowGroups.MailColInx, 1);

                dcnt = 1;
                ITask xtsk = getBpm().findTask(tID);
                if(xtsk == null){continue;}
                IProcessInstance proi = xtsk.getProcessInstance();
                IInformationObject mdoc = (proi != null ? proi.getMainInformationObject() : null);

                IWorkbasket orwb = xtsk.getOriginalWorkbasket();
                String owbn = "";
                //if(orwb != null && orwb.getID() != wbsk.getID()){
                    //owbn = orwb.getFullName();
                //}

                Date tbgn = null, tend = new Date();
                if(xtsk.getReadyDate() != null){
                    tbgn = xtsk.getReadyDate();
                }

                long durd  = 0L;
                double durh  = 0.0;
                if(tend != null && tbgn != null) {
                    long diff = (tend.getTime() > tbgn.getTime() ? tend.getTime() - tbgn.getTime() : tbgn.getTime() - tend.getTime());
                    durd = TimeUnit.DAYS.convert(diff, TimeUnit.MILLISECONDS);
                    durh = ((TimeUnit.MINUTES.convert(diff, TimeUnit.MILLISECONDS) - (durd * 24 * 60)) * 100 / 60) / 100d;
                }

                String rcvf = "", rcvo = "";
                if(xtsk.getPreviousWorkbasket() != null){
                    rcvf = xtsk.getPreviousWorkbasket().getFullName();
                }
                if(tbgn != null){
                    rcvo = (new SimpleDateFormat("dd-MM-yyyy HH:mm")).format(tbgn);
                }

                String prjn = "",  mdno = "", mdrn = "", mdnm = "";
                if(mdoc != null &&  Utils.hasDescriptor((IInformationObject) mdoc, Conf.Descriptors.ProjectNo)){
                    prjn = mdoc.getDescriptorValue(Conf.Descriptors.ProjectNo, String.class);
                }
                if(mdoc != null &&  Utils.hasDescriptor((IInformationObject) mdoc, Conf.Descriptors.DocNumber)){
                    mdno = mdoc.getDescriptorValue(Conf.Descriptors.DocNumber, String.class);
                }
                if(mdoc != null &&  Utils.hasDescriptor((IInformationObject) mdoc, Conf.Descriptors.Revision)){
                    mdrn = mdoc.getDescriptorValue(Conf.Descriptors.Revision, String.class);
                }
                if(mdoc != null &&  Utils.hasDescriptor((IInformationObject) mdoc, Conf.Descriptors.Name)){
                    mdnm = mdoc.getDescriptorValue(Conf.Descriptors.Name, String.class);
                }

                mbms.put("Task" + dcnt, xtsk.getName() + " (" + xtsk.getDescriptorValue("LayerName") + ")");
                mbms.put("ProcessTitle" + dcnt, (proi != null ? proi.getDisplayName() : ""));
                mbms.put("Delegation" + dcnt, (owbn != null ? owbn : ""));
                mbms.put("ProjectNo" + dcnt, (prjn != null  ? prjn : ""));
                mbms.put("DocNo" + dcnt, (mdno != null  ? mdno : ""));
                mbms.put("RevNo" + dcnt, (mdrn != null  ? mdrn : ""));
                mbms.put("DocName" + dcnt, (mdnm != null  ? mdnm : ""));
                mbms.put("ReceivedFrom" + dcnt, (rcvf != null ? rcvf : ""));
                mbms.put("ReceivedOn" + dcnt, (rcvo != null ? rcvo : ""));
                mbms.put("DurDay" + dcnt, durd + "");
                mbms.put("DurHour" + dcnt, durh + "");
                mbms.put("DoxisLink" + dcnt, mcfg.get("webBase") + helper.getTaskURL(xtsk.getID()));
                mbms.put("DoxisLink" + dcnt + ".Text", "( Link )");
            }
            else {
                log.info("Value: {0}", value);
            }

            saveWBInboxExcel(mailExcelPath, Conf.WBInboxMailSheetIndex.Mail, mbms);

            String mailHtmlPath = Utils.convertExcelToHtml(mailExcelPath, Conf.WBInboxMailPaths.EscalationMailPaths + "/" + Conf.MailTemplates.Project + "[" + uniqueId + "].html");
            JSONObject mail = new JSONObject();

            mail.put("To", wbMail);
            //mail.put("Subject","Escalation for Reviewer");
            mail.put("Subject","Escalated Task(s) for Reviewer (" + dcnt + ")");
            mail.put("BodyHTMLFile", mailHtmlPath);

            try {
                //Utils.sendHTMLMail(mcfg, mail);
                Utils.sendHTMLMail(mail, null);
            }catch(Exception ex){
                log.error("EXCP [Send-Mail] : " + ex.getMessage());
            }
        }
        for(String keyStr : DCCList.keySet()) {
            wbMail = keyStr;
            Object value = DCCList.get(keyStr);
            uniqueId = UUID.randomUUID().toString();
            mailExcelPath = Utils.exportDocument(mtpl, Conf.WBInboxMailPaths.EscalationMailPaths, Conf.MailTemplates.Project + "[" + uniqueId + "]");
            int dcnt = 0;
            if (value instanceof JSONArray) {
                JSONArray arr = (JSONArray) value;
                listSize = arr.length();
                loadTableRows(mailExcelPath, Conf.WBInboxMailSheetIndex.Mail, "Task", Conf.WBInboxMailRowGroups.MailColInx, listSize);

                mbms.put("Fullname", keyStr + " (DCC)");
                mbms.put("Count", listSize + "");

                for(int i = 0; i < arr.length(); i++){
                    String tID = arr.getString(i);
                    dcnt++;
                    ITask xtsk = getBpm().findTask(tID);
                    if(xtsk == null){continue;}
                    IProcessInstance proi = xtsk.getProcessInstance();
                    IInformationObject mdoc = (proi != null ? proi.getMainInformationObject() : null);

                    IWorkbasket orwb = xtsk.getOriginalWorkbasket();
                    String owbn = "";
                    //if(orwb != null && orwb.getID() != wbsk.getID()){
                        //owbn = orwb.getFullName();
                    //}

                    Date tbgn = null, tend = new Date();
                    if(xtsk.getReadyDate() != null){
                        tbgn = xtsk.getReadyDate();
                    }

                    long durd  = 0L;
                    double durh  = 0.0;
                    if(tend != null && tbgn != null) {
                        long diff = (tend.getTime() > tbgn.getTime() ? tend.getTime() - tbgn.getTime() : tbgn.getTime() - tend.getTime());
                        durd = TimeUnit.DAYS.convert(diff, TimeUnit.MILLISECONDS);
                        durh = ((TimeUnit.MINUTES.convert(diff, TimeUnit.MILLISECONDS) - (durd * 24 * 60)) * 100 / 60) / 100d;
                    }

                    String rcvf = "", rcvo = "";
                    if(xtsk.getPreviousWorkbasket() != null){
                        rcvf = xtsk.getPreviousWorkbasket().getFullName();
                    }
                    if(tbgn != null){
                        rcvo = (new SimpleDateFormat("dd-MM-yyyy HH:mm")).format(tbgn);
                    }

                    String prjn = "",  mdno = "", mdrn = "", mdnm = "";
                    if(mdoc != null &&  Utils.hasDescriptor((IInformationObject) mdoc, Conf.Descriptors.ProjectNo)){
                        prjn = mdoc.getDescriptorValue(Conf.Descriptors.ProjectNo, String.class);
                    }
                    if(mdoc != null &&  Utils.hasDescriptor((IInformationObject) mdoc, Conf.Descriptors.DocNumber)){
                        mdno = mdoc.getDescriptorValue(Conf.Descriptors.DocNumber, String.class);
                    }
                    if(mdoc != null &&  Utils.hasDescriptor((IInformationObject) mdoc, Conf.Descriptors.Revision)){
                        mdrn = mdoc.getDescriptorValue(Conf.Descriptors.Revision, String.class);
                    }
                    if(mdoc != null &&  Utils.hasDescriptor((IInformationObject) mdoc, Conf.Descriptors.Name)){
                        mdnm = mdoc.getDescriptorValue(Conf.Descriptors.Name, String.class);
                    }

                    mbms.put("Task" + dcnt, xtsk.getName() + " (" + xtsk.getDescriptorValue("LayerName") + ")");
                    mbms.put("ProcessTitle" + dcnt, (proi != null ? proi.getDisplayName() : ""));
                    mbms.put("Delegation" + dcnt, (owbn != null ? owbn : ""));
                    mbms.put("ProjectNo" + dcnt, (prjn != null  ? prjn : ""));
                    mbms.put("DocNo" + dcnt, (mdno != null  ? mdno : ""));
                    mbms.put("RevNo" + dcnt, (mdrn != null  ? mdrn : ""));
                    mbms.put("DocName" + dcnt, (mdnm != null  ? mdnm : ""));
                    mbms.put("ReceivedFrom" + dcnt, (rcvf != null ? rcvf : ""));
                    mbms.put("ReceivedOn" + dcnt, (rcvo != null ? rcvo : ""));
                    mbms.put("DurDay" + dcnt, durd + "");
                    mbms.put("DurHour" + dcnt, durh + "");
                    mbms.put("DoxisLink" + dcnt, mcfg.get("webBase") + helper.getTaskURL(xtsk.getID()));
                    mbms.put("DoxisLink" + dcnt + ".Text", "( Link )");
                }
            }
            else if (value instanceof String) {
                String tID = (String) value;

                mbms.put("Fullname", keyStr + " (DCC)");
                mbms.put("Count", 1 + "");

                loadTableRows(mailExcelPath, Conf.WBInboxMailSheetIndex.Mail, "Task", Conf.WBInboxMailRowGroups.MailColInx, 1);

                dcnt = 1;
                ITask xtsk = getBpm().findTask(tID);
                if(xtsk == null){continue;}
                IProcessInstance proi = xtsk.getProcessInstance();
                IInformationObject mdoc = (proi != null ? proi.getMainInformationObject() : null);

                IWorkbasket orwb = xtsk.getOriginalWorkbasket();
                String owbn = "";
                //if(orwb != null && orwb.getID() != wbsk.getID()){
                    //owbn = orwb.getFullName();
                //}

                Date tbgn = null, tend = new Date();
                if(xtsk.getReadyDate() != null){
                    tbgn = xtsk.getReadyDate();
                }

                long durd  = 0L;
                double durh  = 0.0;
                if(tend != null && tbgn != null) {
                    long diff = (tend.getTime() > tbgn.getTime() ? tend.getTime() - tbgn.getTime() : tbgn.getTime() - tend.getTime());
                    durd = TimeUnit.DAYS.convert(diff, TimeUnit.MILLISECONDS);
                    durh = ((TimeUnit.MINUTES.convert(diff, TimeUnit.MILLISECONDS) - (durd * 24 * 60)) * 100 / 60) / 100d;
                }

                String rcvf = "", rcvo = "";
                if(xtsk.getPreviousWorkbasket() != null){
                    rcvf = xtsk.getPreviousWorkbasket().getFullName();
                }
                if(tbgn != null){
                    rcvo = (new SimpleDateFormat("dd-MM-yyyy HH:mm")).format(tbgn);
                }

                String prjn = "",  mdno = "", mdrn = "", mdnm = "";
                if(mdoc != null &&  Utils.hasDescriptor((IInformationObject) mdoc, Conf.Descriptors.ProjectNo)){
                    prjn = mdoc.getDescriptorValue(Conf.Descriptors.ProjectNo, String.class);
                }
                if(mdoc != null &&  Utils.hasDescriptor((IInformationObject) mdoc, Conf.Descriptors.DocNumber)){
                    mdno = mdoc.getDescriptorValue(Conf.Descriptors.DocNumber, String.class);
                }
                if(mdoc != null &&  Utils.hasDescriptor((IInformationObject) mdoc, Conf.Descriptors.Revision)){
                    mdrn = mdoc.getDescriptorValue(Conf.Descriptors.Revision, String.class);
                }
                if(mdoc != null &&  Utils.hasDescriptor((IInformationObject) mdoc, Conf.Descriptors.Name)){
                    mdnm = mdoc.getDescriptorValue(Conf.Descriptors.Name, String.class);
                }

                mbms.put("Task" + dcnt, xtsk.getName() + " (" + xtsk.getDescriptorValue("LayerName") + ")");
                mbms.put("ProcessTitle" + dcnt, (proi != null ? proi.getDisplayName() : ""));
                mbms.put("Delegation" + dcnt, (owbn != null ? owbn : ""));
                mbms.put("ProjectNo" + dcnt, (prjn != null  ? prjn : ""));
                mbms.put("DocNo" + dcnt, (mdno != null  ? mdno : ""));
                mbms.put("RevNo" + dcnt, (mdrn != null  ? mdrn : ""));
                mbms.put("DocName" + dcnt, (mdnm != null  ? mdnm : ""));
                mbms.put("ReceivedFrom" + dcnt, (rcvf != null ? rcvf : ""));
                mbms.put("ReceivedOn" + dcnt, (rcvo != null ? rcvo : ""));
                mbms.put("DurDay" + dcnt, durd + "");
                mbms.put("DurHour" + dcnt, durh + "");
                mbms.put("DoxisLink" + dcnt, mcfg.get("webBase") + helper.getTaskURL(xtsk.getID()));
                mbms.put("DoxisLink" + dcnt + ".Text", "( Link )");
            }
            else {
                log.info("Value: {0}", value);
            }

            saveWBInboxExcel(mailExcelPath, Conf.WBInboxMailSheetIndex.Mail, mbms);

            String mailHtmlPath = Utils.convertExcelToHtml(mailExcelPath, Conf.WBInboxMailPaths.EscalationMailPaths + "/" + Conf.MailTemplates.Project + "[" + uniqueId + "].html");
            JSONObject mail = new JSONObject();

            mail.put("To", wbMail);
            //mail.put("Subject","Escalation for Reviewer");
            mail.put("Subject","Escalated Task(s) for Reviewer (" + dcnt + ")");
            mail.put("BodyHTMLFile", mailHtmlPath);

            try {
                //Utils.sendHTMLMail(mcfg, mail);
                Utils.sendHTMLMail(mail, null);
            }catch(Exception ex){
                log.error("EXCP [Send-Mail] : " + ex.getMessage());
            }
        }
        log.info("    -> finish ");
    }
    private void notifyForConsalidators(JSONObject mcfg, IDocument mtpl) throws Exception {
        JSONObject MailList = new JSONObject();
        JSONObject DCCList = new JSONObject();
        JSONObject PMList = new JSONObject();
        JSONObject EMList = new JSONObject();
        ArrayList<String> allMails = new ArrayList<>();
        IWorkbasket wbsk = null;
        String wbMail = "", prjMail = "", mailExcelPath = "", uniqueId = "";
        JSONObject mbms = new JSONObject();
        int listSize = 0;

        mainprcss = Utils.getMainReviewProcesses(helper,projects,"Step03");
        for(ITask task : mainprcss){
            log.info("    -> start ");
            int tcnt = 0;

            tcnt++;
            log.info(" *** main review process task [" + tcnt + "] : " + task.getDisplayName());

            String tsid = task.getID();
            log.info(" *** main review process id [" + tcnt + "] : " + tsid);

            String clid = task.getClassID();
            log.info(" *** main review process class id [" + tcnt + "] : " + clid);
            if(!clid.equals(Conf.ClassIDs.ReviewMain)){continue;}

            IProcessInstance proi = task.getProcessInstance();
            if(proi == null){continue;}
            log.info(" *** main review process proi [" + tcnt + "] : " + proi.getDisplayName());

            IDocument wdoc = (IDocument) proi.getMainInformationObject();
            if(wdoc == null){continue;}
            log.info(" *** main review process wdoc [" + tcnt + "] : " + wdoc.getDisplayName());

            String prjn = wdoc.getDescriptorValue(Conf.Descriptors.ProjectNo, String.class);
            if(prjn == null || prjn.isEmpty()){continue;}
            log.info(" *** main review process prjn [" + tcnt + "] : " + prjn);

            IDocument prjCardDoc = getProjectCard(prjn);
            List<String> cnslList = consolidatorList((IInformationObject) task);
            String dccMailList = (prjCardDoc != null ? dccMails(prjCardDoc) : "");
            String pmMail = (prjCardDoc != null ? getPmMail(prjCardDoc) : "");
            String emMail = (prjCardDoc != null ? getEmMail(prjCardDoc) : "");
            prjMail = getMailByPrj(prjCardDoc,"Consalidator");

            String prjDocRvwDrtS = (prjCardDoc != null ? prjCardDoc.getDescriptorValue("ccmPRJCard_DCCDrtn") : "0");

            long prjRvwDrtn = Long.parseLong(prjDocRvwDrtS);

            log.info("    -> class-name : " + task.getName());
            log.info("    -> class-id.doc : " + wdoc.getClassID());
            log.info("    -> class-id.task : " + clid);
            log.info("    -> display : " + wdoc.getDisplayName());

            Date tbgn = null, tend = new Date();
            if(task.getReadyDate() != null){
                tbgn = task.getReadyDate();
            }

            long durd  = 0L;
            double durh  = 0.0;
            if(tend != null && tbgn != null) {
                long diff = (tend.getTime() > tbgn.getTime() ? tend.getTime() - tbgn.getTime() : tbgn.getTime() - tend.getTime());
                durd = TimeUnit.DAYS.convert(diff, TimeUnit.MILLISECONDS);
                if(durd<=prjRvwDrtn){continue;}
                durh = ((TimeUnit.MINUTES.convert(diff, TimeUnit.MILLISECONDS) - (durd * 24 * 60)) * 100 / 60) / 100d;
            }
            ///escalation list creating for consalidators
            MailList.accumulate(prjMail, task.getID());
        }

        for(String keyStr : MailList.keySet()) {
            wbMail = keyStr;
            Object value = MailList.get(keyStr);
            uniqueId = UUID.randomUUID().toString();
            mailExcelPath = Utils.exportDocument(mtpl, Conf.WBInboxMailPaths.EscalationMailPaths, Conf.MailTemplates.Project + "[" + uniqueId + "]");
            int dcnt = 0;
            if (value instanceof JSONArray) {
                JSONArray arr = (JSONArray) value;
                listSize = arr.length();
                loadTableRows(mailExcelPath, Conf.WBInboxMailSheetIndex.Mail, "Task", Conf.WBInboxMailRowGroups.MailColInx, listSize);

                mbms.put("Fullname", keyStr +  " (Prj All mail for Consolidator)");
                mbms.put("Count", listSize + "");

                for(int i = 0; i < arr.length(); i++){
                    String tID = arr.getString(i);
                    dcnt++;
                    ITask xtsk = getBpm().findTask(tID);
                    if(xtsk == null){continue;}
                    IProcessInstance proi = xtsk.getProcessInstance();
                    IInformationObject mdoc = (proi != null ? proi.getMainInformationObject() : null);

                    IWorkbasket orwb = xtsk.getOriginalWorkbasket();
                    String owbn = "";
                    //if(orwb != null && orwb.getID() != wbsk.getID()){
                        //owbn = orwb.getFullName();
                    //}

                    Date tbgn = null, tend = new Date();
                    if(xtsk.getReadyDate() != null){
                        tbgn = xtsk.getReadyDate();
                    }

                    long durd  = 0L;
                    double durh  = 0.0;
                    if(tend != null && tbgn != null) {
                        long diff = (tend.getTime() > tbgn.getTime() ? tend.getTime() - tbgn.getTime() : tbgn.getTime() - tend.getTime());
                        durd = TimeUnit.DAYS.convert(diff, TimeUnit.MILLISECONDS);
                        durh = ((TimeUnit.MINUTES.convert(diff, TimeUnit.MILLISECONDS) - (durd * 24 * 60)) * 100 / 60) / 100d;
                    }

                    String rcvf = "", rcvo = "";
                    if(xtsk.getPreviousWorkbasket() != null){
                        rcvf = xtsk.getPreviousWorkbasket().getFullName();
                    }
                    if(tbgn != null){
                        rcvo = (new SimpleDateFormat("dd-MM-yyyy HH:mm")).format(tbgn);
                    }

                    String prjn = "",  mdno = "", mdrn = "", mdnm = "";
                    if(mdoc != null &&  Utils.hasDescriptor((IInformationObject) mdoc, Conf.Descriptors.ProjectNo)){
                        prjn = mdoc.getDescriptorValue(Conf.Descriptors.ProjectNo, String.class);
                    }
                    if(mdoc != null &&  Utils.hasDescriptor((IInformationObject) mdoc, Conf.Descriptors.DocNumber)){
                        mdno = mdoc.getDescriptorValue(Conf.Descriptors.DocNumber, String.class);
                    }
                    if(mdoc != null &&  Utils.hasDescriptor((IInformationObject) mdoc, Conf.Descriptors.Revision)){
                        mdrn = mdoc.getDescriptorValue(Conf.Descriptors.Revision, String.class);
                    }
                    if(mdoc != null &&  Utils.hasDescriptor((IInformationObject) mdoc, Conf.Descriptors.Name)){
                        mdnm = mdoc.getDescriptorValue(Conf.Descriptors.Name, String.class);
                    }

                    mbms.put("Task" + dcnt, xtsk.getName());
                    mbms.put("ProcessTitle" + dcnt, (proi != null ? proi.getDisplayName() : ""));
                    mbms.put("Delegation" + dcnt, (owbn != null ? owbn : ""));
                    mbms.put("ProjectNo" + dcnt, (prjn != null  ? prjn : ""));
                    mbms.put("DocNo" + dcnt, (mdno != null  ? mdno : ""));
                    mbms.put("RevNo" + dcnt, (mdrn != null  ? mdrn : ""));
                    mbms.put("DocName" + dcnt, (mdnm != null  ? mdnm : ""));
                    mbms.put("ReceivedFrom" + dcnt, (rcvf != null ? rcvf : ""));
                    mbms.put("ReceivedOn" + dcnt, (rcvo != null ? rcvo : ""));
                    mbms.put("DurDay" + dcnt, durd + "");
                    mbms.put("DurHour" + dcnt, durh + "");
                    mbms.put("DoxisLink" + dcnt, mcfg.get("webBase") + helper.getTaskURL(xtsk.getID()));
                    mbms.put("DoxisLink" + dcnt + ".Text", "( Link )");
                }
            }
            else if (value instanceof String) {
                String tID = (String) value;

                mbms.put("Fullname", keyStr +  " (Prj All mail for Consolidator)");
                mbms.put("Count", 1 + "");

                loadTableRows(mailExcelPath, Conf.WBInboxMailSheetIndex.Mail, "Task", Conf.WBInboxMailRowGroups.MailColInx, 1);

                dcnt = 1;
                ITask xtsk = getBpm().findTask(tID);
                if(xtsk == null){continue;}
                IProcessInstance proi = xtsk.getProcessInstance();
                IInformationObject mdoc = (proi != null ? proi.getMainInformationObject() : null);

                IWorkbasket orwb = xtsk.getOriginalWorkbasket();
                String owbn = "";
                //if(orwb != null && orwb.getID() != wbsk.getID()){
                    //owbn = orwb.getFullName();
                //}

                Date tbgn = null, tend = new Date();
                if(xtsk.getReadyDate() != null){
                    tbgn = xtsk.getReadyDate();
                }

                long durd  = 0L;
                double durh  = 0.0;
                if(tend != null && tbgn != null) {
                    long diff = (tend.getTime() > tbgn.getTime() ? tend.getTime() - tbgn.getTime() : tbgn.getTime() - tend.getTime());
                    durd = TimeUnit.DAYS.convert(diff, TimeUnit.MILLISECONDS);
                    durh = ((TimeUnit.MINUTES.convert(diff, TimeUnit.MILLISECONDS) - (durd * 24 * 60)) * 100 / 60) / 100d;
                }

                String rcvf = "", rcvo = "";
                if(xtsk.getPreviousWorkbasket() != null){
                    rcvf = xtsk.getPreviousWorkbasket().getFullName();
                }
                if(tbgn != null){
                    rcvo = (new SimpleDateFormat("dd-MM-yyyy HH:mm")).format(tbgn);
                }

                String prjn = "",  mdno = "", mdrn = "", mdnm = "";
                if(mdoc != null &&  Utils.hasDescriptor((IInformationObject) mdoc, Conf.Descriptors.ProjectNo)){
                    prjn = mdoc.getDescriptorValue(Conf.Descriptors.ProjectNo, String.class);
                }
                if(mdoc != null &&  Utils.hasDescriptor((IInformationObject) mdoc, Conf.Descriptors.DocNumber)){
                    mdno = mdoc.getDescriptorValue(Conf.Descriptors.DocNumber, String.class);
                }
                if(mdoc != null &&  Utils.hasDescriptor((IInformationObject) mdoc, Conf.Descriptors.Revision)){
                    mdrn = mdoc.getDescriptorValue(Conf.Descriptors.Revision, String.class);
                }
                if(mdoc != null &&  Utils.hasDescriptor((IInformationObject) mdoc, Conf.Descriptors.Name)){
                    mdnm = mdoc.getDescriptorValue(Conf.Descriptors.Name, String.class);
                }

                mbms.put("Task" + dcnt, xtsk.getName());
                mbms.put("ProcessTitle" + dcnt, (proi != null ? proi.getDisplayName() : ""));
                mbms.put("Delegation" + dcnt, (owbn != null ? owbn : ""));
                mbms.put("ProjectNo" + dcnt, (prjn != null  ? prjn : ""));
                mbms.put("DocNo" + dcnt, (mdno != null  ? mdno : ""));
                mbms.put("RevNo" + dcnt, (mdrn != null  ? mdrn : ""));
                mbms.put("DocName" + dcnt, (mdnm != null  ? mdnm : ""));
                mbms.put("ReceivedFrom" + dcnt, (rcvf != null ? rcvf : ""));
                mbms.put("ReceivedOn" + dcnt, (rcvo != null ? rcvo : ""));
                mbms.put("DurDay" + dcnt, durd + "");
                mbms.put("DurHour" + dcnt, durh + "");
                mbms.put("DoxisLink" + dcnt, mcfg.get("webBase") + helper.getTaskURL(xtsk.getID()));
                mbms.put("DoxisLink" + dcnt + ".Text", "( Link )");
            }
            else {
                log.info("Value: {0}", value);
            }

            saveWBInboxExcel(mailExcelPath, Conf.WBInboxMailSheetIndex.Mail, mbms);

            String mailHtmlPath = Utils.convertExcelToHtml(mailExcelPath, Conf.WBInboxMailPaths.EscalationMailPaths + "/" + Conf.MailTemplates.Project + "[" + uniqueId + "].html");
            JSONObject mail = new JSONObject();

            mail.put("To", wbMail);
            //mail.put("Subject","Escalation for Consolidator");
            mail.put("Subject","Escalated Task(s) for Consolidator (" + dcnt + ")");
            mail.put("BodyHTMLFile", mailHtmlPath);

            try {
                //Utils.sendHTMLMail(mcfg, mail);
                Utils.sendHTMLMail(mail, null);
            }catch(Exception ex){
                log.error("EXCP [Send-Mail] : " + ex.getMessage());
            }
        }
        log.info("    -> finish ");
    }
    private void notifyForDccs(JSONObject mcfg, IDocument mtpl) throws Exception {
        JSONObject MailList = new JSONObject();
        JSONObject DCCList = new JSONObject();
        JSONObject PMList = new JSONObject();
        JSONObject EMList = new JSONObject();
        ArrayList<String> allMails = new ArrayList<>();
        IWorkbasket wbsk = null;
        String wbMail = "", prjMail = "", mailExcelPath = "", uniqueId = "";
        JSONObject mbms = new JSONObject();
        int listSize = 0;

        mainprcss = Utils.getMainReviewProcesses(helper,projects,"Step04");
        for(ITask task : mainprcss){
            log.info("    -> start ");
            int tcnt = 0;

            tcnt++;
            log.info(" *** main review process task [" + tcnt + "] : " + task.getDisplayName());

            String tsid = task.getID();
            log.info(" *** main review process id [" + tcnt + "] : " + tsid);

            String clid = task.getClassID();
            log.info(" *** main review process class id [" + tcnt + "] : " + clid);
            if(!clid.equals(Conf.ClassIDs.ReviewMain)){continue;}

            IProcessInstance proi = task.getProcessInstance();
            if(proi == null){continue;}
            log.info(" *** main review process proi [" + tcnt + "] : " + proi.getDisplayName());

            IDocument wdoc = (IDocument) proi.getMainInformationObject();
            if(wdoc == null){continue;}
            log.info(" *** main review process wdoc [" + tcnt + "] : " + wdoc.getDisplayName());

            String prjn = wdoc.getDescriptorValue(Conf.Descriptors.ProjectNo, String.class);
            if(prjn == null || prjn.isEmpty()){continue;}
            log.info(" *** main review process prjn [" + tcnt + "] : " + prjn);

            IDocument prjCardDoc = getProjectCard(prjn);
            List<String> cnslList = consolidatorList((IInformationObject) task);
            String dccMailList = (prjCardDoc != null ? dccMails(prjCardDoc) : "");
            String pmMail = (prjCardDoc != null ? getPmMail(prjCardDoc) : "");
            String emMail = (prjCardDoc != null ? getEmMail(prjCardDoc) : "");
            prjMail = getMailByPrj(prjCardDoc,"dcc");

            String prjDocRvwDrtS = (prjCardDoc != null ? prjCardDoc.getDescriptorValue("ccmPRJCard_DCCDrtn") : "0");

            long prjRvwDrtn = Long.parseLong(prjDocRvwDrtS);

            log.info("    -> class-name : " + task.getName());
            log.info("    -> class-id.doc : " + wdoc.getClassID());
            log.info("    -> class-id.task : " + clid);
            log.info("    -> display : " + wdoc.getDisplayName());

            Date tbgn = null, tend = new Date();
            if(task.getReadyDate() != null){
                tbgn = task.getReadyDate();
            }

            long durd  = 0L;
            double durh  = 0.0;
            if(tend != null && tbgn != null) {
                long diff = (tend.getTime() > tbgn.getTime() ? tend.getTime() - tbgn.getTime() : tbgn.getTime() - tend.getTime());
                durd = TimeUnit.DAYS.convert(diff, TimeUnit.MILLISECONDS);
                if(durd<=prjRvwDrtn){continue;}
                durh = ((TimeUnit.MINUTES.convert(diff, TimeUnit.MILLISECONDS) - (durd * 24 * 60)) * 100 / 60) / 100d;
            }

            ///escalation list creating for consalidators
            MailList.accumulate(prjMail, task.getID());

        }

        for(String keyStr : MailList.keySet()) {
            wbMail = keyStr;
            Object value = MailList.get(keyStr);
            uniqueId = UUID.randomUUID().toString();
            mailExcelPath = Utils.exportDocument(mtpl, Conf.WBInboxMailPaths.EscalationMailPaths, Conf.MailTemplates.Project + "[" + uniqueId + "]");
            int dcnt = 0;
            if (value instanceof JSONArray) {
                JSONArray arr = (JSONArray) value;
                listSize = arr.length();
                loadTableRows(mailExcelPath, Conf.WBInboxMailSheetIndex.Mail, "Task", Conf.WBInboxMailRowGroups.MailColInx, listSize);

                mbms.put("Fullname", keyStr +  " (Prj PM and EM mail for Consolidator)");
                mbms.put("Count", listSize + "");

                for(int i = 0; i < arr.length(); i++){
                    String tID = arr.getString(i);
                    dcnt++;
                    ITask xtsk = getBpm().findTask(tID);
                    if(xtsk == null){continue;}
                    IProcessInstance proi = xtsk.getProcessInstance();
                    IInformationObject mdoc = (proi != null ? proi.getMainInformationObject() : null);

                    IWorkbasket orwb = xtsk.getOriginalWorkbasket();
                    String owbn = "";
                    //if(orwb != null && orwb.getID() != wbsk.getID()){
                        //owbn = orwb.getFullName();
                    //}

                    Date tbgn = null, tend = new Date();
                    if(xtsk.getReadyDate() != null){
                        tbgn = xtsk.getReadyDate();
                    }

                    long durd  = 0L;
                    double durh  = 0.0;
                    if(tend != null && tbgn != null) {
                        long diff = (tend.getTime() > tbgn.getTime() ? tend.getTime() - tbgn.getTime() : tbgn.getTime() - tend.getTime());
                        durd = TimeUnit.DAYS.convert(diff, TimeUnit.MILLISECONDS);
                        durh = ((TimeUnit.MINUTES.convert(diff, TimeUnit.MILLISECONDS) - (durd * 24 * 60)) * 100 / 60) / 100d;
                    }

                    String rcvf = "", rcvo = "";
                    if(xtsk.getPreviousWorkbasket() != null){
                        rcvf = xtsk.getPreviousWorkbasket().getFullName();
                    }
                    if(tbgn != null){
                        rcvo = (new SimpleDateFormat("dd-MM-yyyy HH:mm")).format(tbgn);
                    }

                    String prjn = "",  mdno = "", mdrn = "", mdnm = "";
                    if(mdoc != null &&  Utils.hasDescriptor((IInformationObject) mdoc, Conf.Descriptors.ProjectNo)){
                        prjn = mdoc.getDescriptorValue(Conf.Descriptors.ProjectNo, String.class);
                    }
                    if(mdoc != null &&  Utils.hasDescriptor((IInformationObject) mdoc, Conf.Descriptors.DocNumber)){
                        mdno = mdoc.getDescriptorValue(Conf.Descriptors.DocNumber, String.class);
                    }
                    if(mdoc != null &&  Utils.hasDescriptor((IInformationObject) mdoc, Conf.Descriptors.Revision)){
                        mdrn = mdoc.getDescriptorValue(Conf.Descriptors.Revision, String.class);
                    }
                    if(mdoc != null &&  Utils.hasDescriptor((IInformationObject) mdoc, Conf.Descriptors.Name)){
                        mdnm = mdoc.getDescriptorValue(Conf.Descriptors.Name, String.class);
                    }

                    mbms.put("Task" + dcnt, xtsk.getName());
                    mbms.put("ProcessTitle" + dcnt, (proi != null ? proi.getDisplayName() : ""));
                    mbms.put("Delegation" + dcnt, (owbn != null ? owbn : ""));
                    mbms.put("ProjectNo" + dcnt, (prjn != null  ? prjn : ""));
                    mbms.put("DocNo" + dcnt, (mdno != null  ? mdno : ""));
                    mbms.put("RevNo" + dcnt, (mdrn != null  ? mdrn : ""));
                    mbms.put("DocName" + dcnt, (mdnm != null  ? mdnm : ""));
                    mbms.put("ReceivedFrom" + dcnt, (rcvf != null ? rcvf : ""));
                    mbms.put("ReceivedOn" + dcnt, (rcvo != null ? rcvo : ""));
                    mbms.put("DurDay" + dcnt, durd + "");
                    mbms.put("DurHour" + dcnt, durh + "");
                    mbms.put("DoxisLink" + dcnt, mcfg.get("webBase") + helper.getTaskURL(xtsk.getID()));
                    mbms.put("DoxisLink" + dcnt + ".Text", "( Link )");
                }
            }
            else if (value instanceof String) {
                String tID = (String) value;

                mbms.put("Fullname", keyStr +  " (Prj PM and EM mail for Consolidator)");
                mbms.put("Count", 1 + "");

                loadTableRows(mailExcelPath, Conf.WBInboxMailSheetIndex.Mail, "Task", Conf.WBInboxMailRowGroups.MailColInx, 1);

                dcnt = 1;
                ITask xtsk = getBpm().findTask(tID);
                if(xtsk == null){continue;}
                IProcessInstance proi = xtsk.getProcessInstance();
                IInformationObject mdoc = (proi != null ? proi.getMainInformationObject() : null);

                IWorkbasket orwb = xtsk.getOriginalWorkbasket();
                String owbn = "";
                //if(orwb != null && orwb.getID() != wbsk.getID()){
                    //owbn = orwb.getFullName();
                //}

                Date tbgn = null, tend = new Date();
                if(xtsk.getReadyDate() != null){
                    tbgn = xtsk.getReadyDate();
                }

                long durd  = 0L;
                double durh  = 0.0;
                if(tend != null && tbgn != null) {
                    long diff = (tend.getTime() > tbgn.getTime() ? tend.getTime() - tbgn.getTime() : tbgn.getTime() - tend.getTime());
                    durd = TimeUnit.DAYS.convert(diff, TimeUnit.MILLISECONDS);
                    durh = ((TimeUnit.MINUTES.convert(diff, TimeUnit.MILLISECONDS) - (durd * 24 * 60)) * 100 / 60) / 100d;
                }

                String rcvf = "", rcvo = "";
                if(xtsk.getPreviousWorkbasket() != null){
                    rcvf = xtsk.getPreviousWorkbasket().getFullName();
                }
                if(tbgn != null){
                    rcvo = (new SimpleDateFormat("dd-MM-yyyy HH:mm")).format(tbgn);
                }

                String prjn = "",  mdno = "", mdrn = "", mdnm = "";
                if(mdoc != null &&  Utils.hasDescriptor((IInformationObject) mdoc, Conf.Descriptors.ProjectNo)){
                    prjn = mdoc.getDescriptorValue(Conf.Descriptors.ProjectNo, String.class);
                }
                if(mdoc != null &&  Utils.hasDescriptor((IInformationObject) mdoc, Conf.Descriptors.DocNumber)){
                    mdno = mdoc.getDescriptorValue(Conf.Descriptors.DocNumber, String.class);
                }
                if(mdoc != null &&  Utils.hasDescriptor((IInformationObject) mdoc, Conf.Descriptors.Revision)){
                    mdrn = mdoc.getDescriptorValue(Conf.Descriptors.Revision, String.class);
                }
                if(mdoc != null &&  Utils.hasDescriptor((IInformationObject) mdoc, Conf.Descriptors.Name)){
                    mdnm = mdoc.getDescriptorValue(Conf.Descriptors.Name, String.class);
                }

                mbms.put("Task" + dcnt, xtsk.getName());
                mbms.put("ProcessTitle" + dcnt, (proi != null ? proi.getDisplayName() : ""));
                mbms.put("Delegation" + dcnt, (owbn != null ? owbn : ""));
                mbms.put("ProjectNo" + dcnt, (prjn != null  ? prjn : ""));
                mbms.put("DocNo" + dcnt, (mdno != null  ? mdno : ""));
                mbms.put("RevNo" + dcnt, (mdrn != null  ? mdrn : ""));
                mbms.put("DocName" + dcnt, (mdnm != null  ? mdnm : ""));
                mbms.put("ReceivedFrom" + dcnt, (rcvf != null ? rcvf : ""));
                mbms.put("ReceivedOn" + dcnt, (rcvo != null ? rcvo : ""));
                mbms.put("DurDay" + dcnt, durd + "");
                mbms.put("DurHour" + dcnt, durh + "");
                mbms.put("DoxisLink" + dcnt, mcfg.get("webBase") + helper.getTaskURL(xtsk.getID()));
                mbms.put("DoxisLink" + dcnt + ".Text", "( Link )");
            }
            else {
                log.info("Value: {0}", value);
            }

            saveWBInboxExcel(mailExcelPath, Conf.WBInboxMailSheetIndex.Mail, mbms);

            String mailHtmlPath = Utils.convertExcelToHtml(mailExcelPath, Conf.WBInboxMailPaths.EscalationMailPaths + "/" + Conf.MailTemplates.Project + "[" + uniqueId + "].html");
            JSONObject mail = new JSONObject();

            mail.put("To", wbMail);
            //mail.put("Subject","Escalation for DCC");
            mail.put("Subject","Escalated Task(s) for DCC (" + dcnt + ")");
            mail.put("BodyHTMLFile", mailHtmlPath);

            try {
                //Utils.sendHTMLMail(mcfg, mail);
                Utils.sendHTMLMail(mail, null);
            }catch(Exception ex){
                log.error("EXCP [Send-Mail] : " + ex.getMessage());
            }
        }
        log.info("    -> finish ");
    }
    private AgentExecutionResult moveCurrentTaskToNext(ITask eventTask) throws Exception{
        try{
            List<IPossibleDecision> decisions = eventTask.findPossibleDecisions();
            eventTask.setDescriptorValue("Notes","Auto Completed by System (Escalation Raised)");
            eventTask.commit();
            for(IPossibleDecision pdecision : decisions){
                IDecision decision = pdecision.getDecision();
                eventTask.complete(decision);
                break;
            }
        }catch (Exception e){
            log.info("Escalation...error:" + e);
            log.info("Restarting Escalation Agent....");
            return resultRestart("Restarting Agent for Escalation");
        }
        return null;
    }
    public String getUserLoginByWB(String wbID){
        String rtrn = "";
        if(wbID != null) {
            IStringMatrix settingsMatrix = getDocumentServer().getStringMatrixByID("Workbaskets", getSes());
            for (int i = 0; i < settingsMatrix.getRowCount(); i++) {
                String rowID = settingsMatrix.getValue(i, 0);
                if (rowID.equalsIgnoreCase(wbID)) {
                    rtrn = settingsMatrix.getValue(i, 1);
                    break;
                }
            }
        }
        return rtrn;
    }
    public IDocument getProjectCard(String prjNumber)  {
        StringBuilder builder = new StringBuilder();
        builder.append("TYPE = '").append(Conf.ClassIDs.ProjectWorkspace).append("'")
                .append(" AND ")
                .append(Conf.DescriptorLiterals.PrjCardCode).append(" = '").append(prjNumber).append("'");
        String whereClause = builder.toString();
        System.out.println("Where Clause: " + whereClause);

        IInformationObject[] informationObjects = createQuery(new String[]{Conf.Databases.ProjectFolder} , whereClause , 1);
        if(informationObjects.length < 1) {return null;}
        return (IDocument) informationObjects[0];
    }
    public IInformationObject[] createQuery(String[] dbNames , String whereClause , int maxHits){
        String[] databaseNames = dbNames;

        ISerClassFactory fac = getSrv().getClassFactory();
        IQueryParameter que = fac.getQueryParameterInstance(
                getSes() ,
                databaseNames ,
                fac.getExpressionInstance(whereClause) ,
                null,null);
        if(maxHits > 0) {
            que.setMaxHits(maxHits);
            que.setHitLimit(maxHits + 1);
            que.setHitLimitThreshold(maxHits + 1);
        }
        IDocumentHitList hits = que.getSession() != null? que.getSession().getDocumentServer().query(que, que.getSession()):null;
        if(hits == null) return null;
        else return hits.getInformationObjects();
    }
    public String dccMails(IDocument prjCard){
        String[] membersIDs = new String[0];
        String rtrn = "";
        List<String> dcc = new ArrayList<>();
        String dccMembers = prjCard.getDescriptorValue("ccmPrjCard_DccList");
        membersIDs = dccMembers.replace("[", "").replace("]", "").split(",");
        for(String mmbrKey : Arrays.asList(membersIDs)){
            IUser mmbr = getDocumentServer().getUserByLoginName(getSes() , getUserLoginByWB(mmbrKey));
            if(mmbr == null){continue;}
            dcc.add(mmbr.getEMailAddress());
        }
        rtrn = String.join(",",dcc);
        return rtrn;
    }
    public String getPmMail(IDocument prjCard){
        String rtrn = "";
        String pmMember = prjCard.getDescriptorValue("ccmPRJCard_prjmngr");
        IUser mmbr = getDocumentServer().getUserByLoginName(getSes() , getUserLoginByWB(pmMember));
        if(mmbr != null){rtrn = mmbr.getEMailAddress();}
        return rtrn;
    }
    public String getEmMail(IDocument prjCard){
        String rtrn = "";
        String pmMember = prjCard.getDescriptorValue("ccmPRJCard_EngMng");
        IUser mmbr = getDocumentServer().getUserByLoginName(getSes() , getUserLoginByWB(pmMember));
        if(mmbr != null){rtrn = mmbr.getEMailAddress();}
        return rtrn;
    }
    public String getMailByPrj(IDocument prjCard,String type){
        String rtrn = "";
        ArrayList<String> allMails = new ArrayList<>();
        String dccMailList = (prjCard != null ? dccMails(prjCard) : "");
        String pmMail = (prjCard != null ? getPmMail(prjCard) : "");
        String emMail = (prjCard != null ? getEmMail(prjCard) : "");
        if(type == "Consalidator"){allMails.add(dccMailList);}
        allMails.add(pmMail);
        allMails.add(emMail);
        rtrn = String.join(",", allMails);
        return rtrn;
    }
    public List<String> consolidatorList(IInformationObject doc){
        String[] membersIDs = new String[0];
        String members = (doc.getDescriptorValue("ccmConsolidatorList") == null ? "" : doc.getDescriptorValue("ccmConsolidatorList"));
        membersIDs = members.replace("[", "").replace("]", "").split(",");
        List<String> rtrn = new ArrayList<>(Arrays.asList(membersIDs));
        return rtrn;
    }
    public String getNumber(String nrName){
        String rtrn = "";
        String pattern = "";
        String nrStartS = "";
        String nrCurrentS = "";
        IStringMatrix settingsMatrix = getDocumentServer().getStringMatrixByID("SYS_NUMBER_RANGES", getSes());
        for (int i = 0; i < settingsMatrix.getRowCount(); i++) {
            String rowID = settingsMatrix.getValue(i, 0);
            if (rowID.equalsIgnoreCase(nrName)) {
                pattern = settingsMatrix.getValue(i, 2);
                nrStartS = settingsMatrix.getValue(i, 3);
                nrCurrentS = settingsMatrix.getValue(i, 5);
                break;
            }
        }


        return rtrn;
    }
}