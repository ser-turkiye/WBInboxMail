package ser;

import java.util.List;

public class Conf {

    public static class Licences{
        public static final String SPIRE_XLS = "XF0BAJOfSVJ6sw8TuTirgf9bo3BKkyPAbKRGLghup0t41I4JxXOS9D0G0AoHnfxP2lRRCH42qefAkbhn1e1X82GnwfVezKuihOE3lVhhpq1ZIYQ7W6KE51Xg6LES0p1Mwdzlc3hb3pPc9jF5BsB/A18RQ1CSz4cwpBd3NDMxEnTCOaWNtUMzOZDO7AhCsHcIEwyj67WdyVOUjuTeDftD7mHsnjWyPTpRw+CZHFLkio8rRIaOOJ8ZGedeQU2a4sn8+NeYy73a6Konq5xBZcEzAdQo05Z8EuuYtW5/DSexdj8vqqyfn0OtEz6iYiQuUaS+/qOSivjAUKAuzbVlq2U/x7muOtHDJO+U56kIbLk24whEhekV19ptLmToWbNelIixCTesrqB17Dv9dXyozKzBln8GIYDfRuwtfiFWMn/VPAPUBiyB0FRzUWE2rMTQEWtgDYH6kakoqQcJi/zyHCx+Ey9mLsk3+n3S/hTBOyTXFDGpXromWz6BiQHgE5OicVcgrG6NCXeBJRhFIhi8P+MI4+oxgcMN+2kDp4YJt2MfQNYKjbugJAEAovJorm/flgY4kn9TaBnxVUZgzX5zudw434j5UZpj1kCQew65rYiMoQYtsw+o6NKT1dZvKYM0f1GtlHFXIcefihEMVGcbhkV3qex4GxQWMt8Ubtn2Nvled8gTpR/SPb3Ui3aJuIuFt7VQZ/FfMglpIAjkLXayHr/466i9YmQjkyD/Fhk2GYuRTMvbbYl+7NC4RMhUT0ny7a4nyyHc9wv7rhyvduw/xInGbiiWpio7KyRDeTyb2LPP1w8f/FaiFBxyk3/NmJq6dmZ8PeV3lA+GVV7mxGf1prdD5tMivKawFsR05M4bwIp7eXt7htqEICig9AuqezADMiAL1uVtne0m4CZ81asp4Hn9UXem4WDWXN4JdOsCI4E91iPATO/zRUWpk5R6VsSJspMNJhjmUumZNkGU9SjBv1QMn435HGTM5u2eC1U4EcIuU/YlPPSvCa34PRMqH201FWoOJNYP0UE4wMUJSsEmRIbhiulRISyNZqF/FezfcYIO4ujqLmGmdRD8yrnKyeTDFxEIfhXSP/tj2mZ3h89iYonhry/e8lef+N1EwRU1S8JSGeZVND8luW+3tVTkV6Er65FyFh33vvTebdGkMS2236rHtJ9WO1zOEhXjcTly2KrwrLhnIHWoVguCVOMkuWU2HIU+BqVcgR3PXWD0eS2KkD9HumdTVChndpUMaP1yfxhNjsGb/k08n3kBguG7ZOsBe5j/GxJBkvBZae90lvzDYGyKxyHA/dLqv402N5BjFsUZzGeqWXInItPsDe7aDKvk7S25johx79pcLsd3maZo2K9bJEuiT4lGf1XZoOVKoy7u8X+1BsHK8/zdmN0+uNG+PE/T8EdTTRH/CQcyWMiv/BAC+QSRIE52VQr4KH5Zfo31V4wH+jg2Qe03muVK6rr9JmwE0mokFWQEXgUmchxa5HbA1H82ro7cUVFjWLArHr5oYkHhPJYiKZXroFm/CJuWBY/NWVTvBHB3YZuLkDr2J9v29HymoTOg8WZ8i02DGgOum66/AC2HJcZCPSQ591CytBbsINhziWPqvwS4XMFsoL6OyNGqQLtXIdb3qYEd5bvgaAs=";

    }
    public static class MailTemplates{
        public static final String Project = "MAIL_WORKBASKET_LIST";
    }

    public static class WBInboxMailSheetIndex {
        public static final Integer Mail = 0;
    }
    public static class WBInboxMailRowGroups {
        public static final List<Integer> MailHideCols = List.of(0);
        public static final Integer MailColInx = 0;
    }
    public static class WBInboxMailPaths {
        public static final String MainPath = "C:/tmp2/bulk/wb-inboxmails";
        public static final String WebBase = "http://localhost/webcube/";
    }
    public static class MainWFUpdateSheetIndex {
        public static final Integer Mail = 0;
    }
    public static class Databases{
        public static final String Company = "D_QCON";
        public static final String ProjectFolder = "PRJ_FOLDER";
    }
    public static class ClassIDs{
        public static final String Template = "b9cf43d1-a4d3-482f-9806-44ae64c6139d";
        public static final String Transmittal = "8bf0a09b-b569-4aef-984b-78cf1644ca19";
        public static final String SubReview = "629a28c4-6c36-44d0-90f7-1e5802f038e8";
        public static final String ReviewMain = "69d42aaf-6978-4b5a-8178-88a78f4b3158";
        public static final String ProjectWorkspace = "32e74338-d268-484d-99b0-f90187240549";
    }
    public static class Descriptors{
        public static final String ProjectNo = "ccmPRJCard_code";
        public static final String DocNumber = "ccmPrjDocNumber";
        public static final String Revision = "ccmPrjDocRevision";
        public static final String Name = "ObjectName";
        public static final String TemplateName = "ObjectNumberExternal";

    }
    public static class DescriptorLiterals{
        public static final String PrjCardCode = "CCMPRJCARD_CODE";
        public static final String ObjectNumberExternal = "OBJECTNUMBER2";
    }
}
