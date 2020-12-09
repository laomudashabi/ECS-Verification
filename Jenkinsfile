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
        bat 'echo The target path to store AD projects is: %TargetPath%'
      }
    }

    stage('Call AD COM Server') {
      steps {
        bat(script: 'py -2.7 API_AddLib.py %TestScope% %TargetProject% %TargetPath%', returnStatus: true, returnStdout: true)
      }
    }

  }
}