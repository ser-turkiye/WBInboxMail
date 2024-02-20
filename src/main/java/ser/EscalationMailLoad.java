package ser;

import com.ser.blueline.*;
import com.ser.blueline.bpm.IProcessInstance;
import com.ser.blueline.bpm.ITask;
import com.ser.blueline.bpm.IWorkbasket;
import com.ser.blueline.bpm.IWorkbasketContent;
import com.ser.blueline.metaDataComponents.IStringMatrix;
import de.ser.doxis4.agentserver.UnifiedAgent;
import org.apache.commons.collections4.IteratorUtils;
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
    JSONObject DCCList = new JSONObject();
    JSONObject ConsalidatorList = new JSONObject();
    JSONObject PMList = new JSONObject();
    JSONObject tasks = new JSONObject();
    List<ITask> subprcss = new ArrayList<>();
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
            helper = new ProcessHelper(Utils.session);
            JSONObject mcfg = Utils.getMailConfig();

            projects = Utils.getProjectWorkspaces(helper);
            mailTemplate = null;

            for(String prjn : projects.keySet()){
                IInformationObject prjt = (IInformationObject) projects.get(prjn);
                IDocument dtpl = Utils.getTemplateDocument(prjt, Conf.MailTemplates.Project);
                if(dtpl == null){continue;}
                mailTemplate = dtpl;
            }

            notifyForReviewer(mcfg,mailTemplate);

            if(mailTemplate == null){throw new Exception("Mail template not found.");}

           /* List<IWorkbasket> wbs = Utils.bpm.getWorkbaskets();
            for (IWorkbasket wb : wbs){
                IWorkbasket swb = Utils.bpm.getWorkbasket(wb.getID());
                if(swb.getNotifyEMail() == null){continue;}
                if(!swb.getNotifyEMail().contains("bulent")){continue;}
                runWorkbasketForReviewer(swb, mcfg, mailTemplate);
            }*/
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
        JSONObject docs = new JSONObject();
        IWorkbasket wbsk = null;
        String wbMail = "", mailExcelPath = "";
        JSONObject mbms = new JSONObject();

        subprcss = Utils.getSubReviewProcesses(helper,projects);
        for(ITask task : subprcss){
            log.info("    -> start ");
            int tcnt = 0;

            tcnt++;
            log.info(" *** subprocess task [" + tcnt + "] : " + task.getDisplayName());

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

            if(docs.has(task.getID())){continue;}

            List<String> cnslList = consolidatorList((IInformationObject) task);

            IDocument prjCardDoc = getProjectCard(prjn);
            String prjDocRvwDrtS = (prjCardDoc != null ? prjCardDoc.getDescriptorValue("ccmPRJCard_ReviewerDrtn") : "0");

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

            ///escalation list creating by consalidators
            for(String cnslId : cnslList) {
                if(cnslId.isEmpty()){continue;}
                IUser mmbr = getDocumentServer().getUserByLoginName(getSes() , getUserLoginByWB(cnslId));
                if(mmbr == null){continue;}
                ConsalidatorList.accumulate(mmbr.getEMailAddress(), task.getID());
                docs.put(task.getID(), task);
            }
        }

        String uniqueId = "";
        int listSize = 0;
        for(String keyStr : ConsalidatorList.keySet()) {
            Object value = ConsalidatorList.get(keyStr);
            uniqueId = UUID.randomUUID().toString();
            mailExcelPath = Utils.exportDocument(mtpl, Conf.WBInboxMailPaths.EscalationMailPaths, Conf.MailTemplates.Project + "[" + uniqueId + "]");
            int dcnt = 0;
            if (value instanceof JSONObject) {

            } else if (value instanceof JSONArray) {
                JSONArray arr = (JSONArray) value;
                listSize = arr.length();
                loadTableRows(mailExcelPath, Conf.WBInboxMailSheetIndex.Mail, "Task", Conf.WBInboxMailRowGroups.MailColInx, arr.length());

                mbms.put("Fullname", keyStr);
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
                    if(orwb != null && orwb.getID() != wbsk.getID()){
                        owbn = orwb.getFullName();
                    }

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

                mbms.put("Fullname", keyStr);
                mbms.put("Count", 1 + "");

                loadTableRows(mailExcelPath, Conf.WBInboxMailSheetIndex.Mail, "Task", Conf.WBInboxMailRowGroups.MailColInx, 1);

                dcnt = 1;
                ITask xtsk = getBpm().findTask(tID);
                if(xtsk == null){continue;}
                IProcessInstance proi = xtsk.getProcessInstance();
                IInformationObject mdoc = (proi != null ? proi.getMainInformationObject() : null);

                IWorkbasket orwb = xtsk.getOriginalWorkbasket();
                String owbn = "";
                if(orwb != null && orwb.getID() != wbsk.getID()){
                    owbn = orwb.getFullName();
                }

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

            String mailHtmlPath = Utils.convertExcelToHtml(mailExcelPath,
                    Conf.WBInboxMailPaths.EscalationMailPaths + "/" + Conf.MailTemplates.Project + "[" + uniqueId + "].html");
            JSONObject mail = new JSONObject();

            mail.put("To", wbMail);
            mail.put("Subject",
                    "Workbox Notification ({Count}) Task is waiting your action"
                            .replace("{Count}", listSize + "")
            );
            mail.put("BodyHTMLFile", mailHtmlPath);

            try {
                Utils.sendHTMLMail(mcfg, mail);
            }catch(Exception ex){
                log.error("EXCP [Send-Mail] : " + ex.getMessage());
            }
        }
        log.info("    -> finish ");
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
    public List<String> dccList(IDocument prjCard){
        String[] membersIDs = new String[0];
        String dccMembers = prjCard.getDescriptorValue("ccmPrjCard_DccList");
        membersIDs = dccMembers.replace("[", "").replace("]", "").split(",");
        List<String> rtrn = new ArrayList<>(Arrays.asList(membersIDs));
        return rtrn;
    }
    public List<String> consolidatorList(IInformationObject doc){
        String[] membersIDs = new String[0];
        String members = (doc.getDescriptorValue("ccmConsolidatorList") == null ? "" : doc.getDescriptorValue("ccmConsolidatorList"));
        membersIDs = members.replace("[", "").replace("]", "").split(",");
        List<String> rtrn = new ArrayList<>(Arrays.asList(membersIDs));
        return rtrn;
    }
}