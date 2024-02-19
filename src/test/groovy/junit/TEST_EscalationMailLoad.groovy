package junit

import de.ser.doxis4.agentserver.AgentExecutionResult
import org.junit.*
import ser.EscalationMailLoad
import ser.WBInboxMailLoad

class TEST_EscalationMailLoad {

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
        def agent = new EscalationMailLoad();

        binding["AGENT_EVENT_OBJECT_CLIENT_ID"] = "ST03BPM24d3523a79-f208-46b6-99ef-4a1489102fcd182024-02-19T08:16:58.072Z010"

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
