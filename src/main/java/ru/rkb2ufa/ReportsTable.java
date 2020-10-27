package ru.rkb2ufa;

import java.io.File;
import static ru.rkb2ufa.MainServlet.REPORTS_DIR;

public class ReportsTable {

    public ReportsTable(String FullFileName, int id) {
        this.orgName = FullFileName.substring(FullFileName.lastIndexOf(File.separator)+1, FullFileName.indexOf("(")-1);
        this.tipOtcheta = FullFileName.substring(FullFileName.indexOf("(")+1, FullFileName.indexOf(")"));
        this.period = FullFileName.substring(FullFileName.indexOf("[")+1, FullFileName.indexOf("]"));
        //this.dataVremya = FullFileName.substring(FullFileName.indexOf("]")+2, FullFileName.lastIndexOf("."));
        //this.downloadLink = "<a href=\"."+File.separator + FullFileName.substring(FullFileName.indexOf(REPORTS_DIR), FullFileName.length())+"\" >скач.</a>";
        //this.removeLink = "<a href=\"delete?id="+id+"\" onclick=\"window.location = 'list'\" >удал.</a>";
        this.dataVremya = FullFileName.substring(FullFileName.indexOf("]")+2, FullFileName.lastIndexOf(".")).replace("__", " ").replace('-',':');
        this.downloadLink = "<a href=\"."+File.separator + FullFileName.substring(FullFileName.indexOf(REPORTS_DIR), FullFileName.length())+"\" download=\"\"><button>Cкачать</button></a>";
        this.removeLink = "<a href=\"delete?id="+id+"\" onclick=\"window.location = 'list'\" ><button>Удалить</button></a>";
    }

    public void setOrgName(String orgName) { this.orgName = orgName; }

    public void setTipOtcheta(String tipOtcheta) {
        this.tipOtcheta = tipOtcheta;
    }

    public void setPeriod(String period) {
        this.period = period;
    }

    public void setDataVremya(String DataVremya){ this.dataVremya = DataVremya;  }

    public void setDownloadLink(String DownloadLink) { this.downloadLink = DownloadLink;  }

    public void setRemoveLink(String RemoveLink) { this.removeLink = RemoveLink; }

    public String getOrgName() {
        return orgName;
    }

    public String getTipOtcheta() {
        return tipOtcheta;
    }

    public String getPeriod() {
        return period;
    }

    public String getDataVremya() {
        return dataVremya;
    }

    public String getDownloadLink() {
        return downloadLink;
    }

    public String getRemoveLink() {
        return removeLink;
    }

    String orgName;
    String tipOtcheta;
    String period;
    String dataVremya;
    String downloadLink;
    String removeLink;
}
