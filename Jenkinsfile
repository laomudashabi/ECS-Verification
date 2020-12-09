pipeline {
  agent {
    node {
      label 'ECS_Verification'
    }

  }
  stages {
    stage('Call AD COM Server') {
      steps {
        bat(script: 'py -2.7 API_AddLib.py %TestScope% "C:\\\\Program Files (x86)\\\\CVDE-Interface\\\\Template_TestExecutionProject_AD5-6.zip" "C:\\\\Users\\\\a269028\\\\Desktop\\\\"', returnStatus: true, returnStdout: true)
      }
    }

    stage('Print out') {
      steps {
        bat 'echo %TestScope%'
      }
    }

  }
}