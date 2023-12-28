package junit

import de.ser.doxis4.agentserver.AgentExecutionResult
import org.junit.*
import ser.WBInboxMailTest

class TEST_WBInboxMailTest {

    Binding binding

    @BeforeClass
    static void initSessionPool() {
        AgentTester.initSessionPool()
    }

    @Before
    void retrieveBinding() {
        binding = AgentTester.retrieveBinding()
    }

    @Test
    void testForAgentResult() {
        def agent = new WBInboxMailTest();

        binding["AGENT_EVENT_OBJECT_CLIENT_ID"] = "ST03BPM248966319a-f756-48c3-9eb5-9c2c8898ee30182023-11-08T13:54:59.194Z016"

        def result = (AgentExecutionResult) agent.execute(binding.variables)
        assert result.resultCode == 0
    }

    @Test
    void testForJavaAgentMethod() {
        //def agent = new JavaAgent()
        //agent.initializeGroovyBlueline(binding.variables)
        //assert agent.getServerVersion().contains("Linux")
    }

    @After
    void releaseBinding() {
        AgentTester.releaseBinding(binding)
    }

    @AfterClass
    static void closeSessionPool() {
        AgentTester.closeSessionPool()
    }
}
