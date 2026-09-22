pipeline {
    agent any

    stages {

        stage('Install Dependencies') {
            steps {
                script {
                    if (isUnix()) {
                        sh 'python3 -m pip install -r requirements.txt'
                    } else {
                        bat 'py -m pip install -r requirements.txt'
                    }
                }
            }
        }

        stage('Run Monitoring') {
            steps {
                script {
                    if (isUnix()) {
                        sh 'python3 monitoring.py'
                    } else {
                        bat 'py monitoring.py'
                    }
                }
            }
        }
    }

    post {
        always {
            archiveArtifacts artifacts: 'monitoring_output/*.xlsx',
                             allowEmptyArchive: false,
                             fingerprint: true
        }
    }
}
