package junit

import de.ser.doxis4.agentserver.AgentExecutionResult
import org.junit.*
import ser.*

class TEST_WBInboxMailLoad {

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
        def agent = new WBInboxMailLoad();

        binding["AGENT_EVENT_OBJECT_CLIENT_ID"] = "ST03BPM24e8763f96-7b9b-4aff-9da8-2551c3731860182023-11-10T12:04:17.822Z014"

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
