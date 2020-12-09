pipeline {
  agent {
    node {
      label 'ECS_Verification'
    }

  }
  stages {
    stage('Call AD COM Server') {
      steps {
        bat(script: 'py -2.7 API_AddLib.py %TestScope% %TargetProject% %TargetPath%', returnStatus: true, returnStdout: true)
      }
    }

    stage('Print out') {
      steps {
        bat 'echo %TestScope%'
      }
    }

  }
}
