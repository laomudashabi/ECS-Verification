pipeline {
  agent {
    node {
      label 'ECS_Verification'
    }

  }
  stages {
    stage('Print out') {
      steps {
        bat 'echo TestScope directory is: %TestScope%'
        bat 'echo The AD project loaded is: %TargetProject%'
        bat 'echo The software package path is: %TargetPath%'
        powershell 'echo The nodes under test are: %Nodes%'
      }
    }

    stage('Call AD COM Server') {
      steps {
        bat(script: 'py -2.7 ExecTestScope.py %TestScope% %TargetProject% %SWCollection% %TestObject% %Nodes%', returnStatus: true, returnStdout: true)
      }
    }

  }
}